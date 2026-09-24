//! Scalar admission and fixed-state kernels for the core statistical reducers.
//!
//! The resolver-aware value evaluator uses the same [`StatisticalKernel`] as
//! the scalar evaluator.  The kernel deliberately knows nothing about cell
//! reads or formula-error ordering: callers classify each value according to
//! the function's sequence signature, retain source-order formula errors, and
//! then feed only the admitted numeric values here.  Its state is bounded and
//! allocation-free.  Numeric sums and averages, counts, and extrema reuse the
//! exact fixed-width implementation in [`super::numerics`].

use super::numerics::{NumericAggregate, NumericOperation};
use super::{EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, WorkingValue};

/// The seventeen statistical reducers implemented by the statistical evaluator.
///
/// This type is private to the formula evaluator and its resolver-aware child
/// modules.  Keeping the function identity here gives both profiles one
/// source for numeric-operation selection and empty-result behavior without
/// adding a public API surface.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum StatisticalFunction {
    Count,
    CountA,
    CountBlank,
    Average,
    AverageA,
    Minimum,
    Maximum,
    MinimumA,
    MaximumA,
    Variance,
    VarianceA,
    VarianceP,
    VariancePA,
    StandardDeviation,
    StandardDeviationA,
    StandardDeviationP,
    StandardDeviationPA,
}

impl StatisticalFunction {
    /// Resolve one case-insensitive OpenFormula function name.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        Some(if name.eq_ignore_ascii_case("COUNT") {
            Self::Count
        } else if name.eq_ignore_ascii_case("COUNTA") {
            Self::CountA
        } else if name.eq_ignore_ascii_case("COUNTBLANK") {
            Self::CountBlank
        } else if name.eq_ignore_ascii_case("AVERAGE") {
            Self::Average
        } else if name.eq_ignore_ascii_case("AVERAGEA") {
            Self::AverageA
        } else if name.eq_ignore_ascii_case("MIN") {
            Self::Minimum
        } else if name.eq_ignore_ascii_case("MAX") {
            Self::Maximum
        } else if name.eq_ignore_ascii_case("MINA") {
            Self::MinimumA
        } else if name.eq_ignore_ascii_case("MAXA") {
            Self::MaximumA
        } else if name.eq_ignore_ascii_case("VAR") {
            Self::Variance
        } else if name.eq_ignore_ascii_case("VARA") {
            Self::VarianceA
        } else if name.eq_ignore_ascii_case("VARP") {
            Self::VarianceP
        } else if name.eq_ignore_ascii_case("VARPA") {
            Self::VariancePA
        } else if name.eq_ignore_ascii_case("STDEV") {
            Self::StandardDeviation
        } else if name.eq_ignore_ascii_case("STDEVA") {
            Self::StandardDeviationA
        } else if name.eq_ignore_ascii_case("STDEVP") {
            Self::StandardDeviationP
        } else if name.eq_ignore_ascii_case("STDEVPA") {
            Self::StandardDeviationPA
        } else {
            return None;
        })
    }

    pub(super) const fn numeric_operation(self) -> NumericOperation {
        match self {
            Self::Count => NumericOperation::Count,
            Self::Average | Self::AverageA => NumericOperation::Average,
            Self::Minimum | Self::MinimumA => NumericOperation::Minimum,
            Self::Maximum | Self::MaximumA => NumericOperation::Maximum,
            Self::CountA | Self::CountBlank => NumericOperation::Count,
            Self::Variance | Self::VarianceA => NumericOperation::SampleVariance,
            Self::VarianceP | Self::VariancePA => NumericOperation::PopulationVariance,
            Self::StandardDeviation | Self::StandardDeviationA => {
                NumericOperation::SampleStandardDeviation
            },
            Self::StandardDeviationP | Self::StandardDeviationPA => {
                NumericOperation::PopulationStandardDeviation
            },
        }
    }

    pub(super) const fn is_a(self) -> bool {
        matches!(
            self,
            Self::AverageA
                | Self::MinimumA
                | Self::MaximumA
                | Self::VarianceA
                | Self::VariancePA
                | Self::StandardDeviationA
                | Self::StandardDeviationPA
        )
    }

    pub(super) const fn is_count(self) -> bool {
        matches!(self, Self::Count)
    }

    pub(super) const fn is_counta(self) -> bool {
        matches!(self, Self::CountA)
    }

    pub(super) const fn is_countblank(self) -> bool {
        matches!(self, Self::CountBlank)
    }

    pub(super) const fn includes_text(self) -> bool {
        self.is_a()
    }

    pub(super) const fn empty_error(self) -> Option<ScalarError> {
        match self {
            Self::Average | Self::AverageA => Some(ScalarError::DivisionByZero),
            Self::Variance
            | Self::VarianceA
            | Self::VarianceP
            | Self::VariancePA
            | Self::StandardDeviation
            | Self::StandardDeviationA
            | Self::StandardDeviationP
            | Self::StandardDeviationPA => Some(ScalarError::Value),
            _ => None,
        }
    }

    pub(super) const fn zero_when_empty(self) -> bool {
        matches!(
            self,
            Self::Minimum | Self::Maximum | Self::MinimumA | Self::MaximumA
        )
    }
}

/// Return whether `name` is one of the core statistical reducers.
pub(super) fn is_statistical_function(name: &str) -> bool {
    StatisticalFunction::from_name(name).is_some()
}

/// Fixed-state statistical accumulator shared with the value evaluator.
///
/// `push_number` is used for admitted numbers and for A-variant conversions
/// (Text to zero and Logical to zero or one). `push_nonblank` is used by
/// COUNTA, while `push_blank` is used by COUNTBLANK.  Formula errors are
/// intentionally retained by each caller so source position and the distinct
/// COUNT/COUNTA error rules remain visible at the scan boundary.
#[derive(Clone, Copy, Debug)]
pub(super) struct StatisticalKernel {
    function: StatisticalFunction,
    numeric: NumericAggregate,
    count: u64,
}

impl StatisticalKernel {
    /// Create the fixed state for one statistical function.
    pub(super) fn new(function: StatisticalFunction) -> Self {
        Self {
            function,
            numeric: NumericAggregate::new(function.numeric_operation()),
            count: 0,
        }
    }

    /// Add one admitted finite Number.
    pub(super) fn push_number(&mut self, value: f64) -> Result<(), ScalarError> {
        if matches!(
            self.function,
            StatisticalFunction::CountA | StatisticalFunction::CountBlank
        ) {
            return Err(ScalarError::Value);
        }
        self.numeric.push_number(value)
    }

    /// Count one non-empty value for COUNTA.
    pub(super) fn push_nonblank(&mut self) -> Result<(), ScalarError> {
        if self.function != StatisticalFunction::CountA {
            return Err(ScalarError::Value);
        }
        self.count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        Ok(())
    }

    /// Count one blank cell for COUNTBLANK.
    pub(super) fn push_blank(&mut self) -> Result<(), ScalarError> {
        if self.function != StatisticalFunction::CountBlank {
            return Err(ScalarError::Value);
        }
        self.count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        Ok(())
    }

    /// Return the number of admitted numeric values.
    pub(super) const fn numeric_count(&self) -> u64 {
        self.numeric.count()
    }

    /// Finish according to the function's empty-selection profile.
    pub(super) fn finish(&self) -> Result<f64, ScalarError> {
        if self.function == StatisticalFunction::CountA
            || self.function == StatisticalFunction::CountBlank
        {
            return Ok(self.count as f64);
        }
        let result = self.numeric.result()?;
        if result.is_none() {
            if let Some(error) = self.function.empty_error() {
                return Err(error);
            }
        }
        match result {
            Some(value) => {
                if self.function.zero_when_empty() && value == 0.0 {
                    Ok(0.0)
                } else {
                    Ok(value)
                }
            },
            // MIN/MAX and their A variants use the specified zero identity
            // when no Number was admitted.
            None => Ok(0.0),
        }
    }
}

/// Apply one statistical reducer after eager scalar argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = StatisticalFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::UnsupportedKind::Function,
        ));
    };

    let count = node.child_count();
    if function == StatisticalFunction::CountBlank {
        if count != 1 {
            return evaluator.finish_invalid_arity(node);
        }
        // COUNTBLANK's sole parameter is a ReferenceList.  A scalar value is
        // a well-formed expression but not an admissible parameter; a
        // reference never reaches this scalar profile because the visitor
        // reports its capability refusal before dispatch.
        reverse_value_tail(evaluator, count)?;
        let value = evaluator.pop_value()?;
        return match value {
            WorkingValue::Error(error) => evaluator.push_value(WorkingValue::Error(error)),
            WorkingValue::Number(_)
            | WorkingValue::Logical(_)
            | WorkingValue::Text(_)
            | WorkingValue::Complex(_) => {
                evaluator.push_value(WorkingValue::Error(ScalarError::Value))
            },
        };
    }

    if count == 0 {
        if let Some(error) = function.empty_error() {
            return evaluator.push_value(WorkingValue::Error(error));
        }
        // COUNT/COUNTA and MIN/MAX have a useful zero identity in this
        // bounded profile.  The ODF text permits an Error or zero for the
        // count functions and recommends zero for MIN; zero keeps omitted
        // variadic calls composable and matches the aggregate profile.
        return evaluator.push_value(WorkingValue::Number(0.0));
    }

    reverse_value_tail(evaluator, count)?;
    let mut kernel = StatisticalKernel::new(function);
    let mut formula_error = None;
    let mut generated_error = None;

    for _ in 0..count {
        evaluator.charge_work(1)?;
        let value = evaluator.pop_value()?;
        observe_scalar(
            evaluator,
            &mut kernel,
            function,
            value,
            &mut formula_error,
            &mut generated_error,
        )?;
    }

    let value = match formula_error.or(generated_error) {
        Some(error) => WorkingValue::Error(error),
        None => {
            evaluator.charge_work(1)?;
            match kernel.finish() {
                Ok(value) => WorkingValue::Number(value),
                Err(error) => WorkingValue::Error(error),
            }
        },
    };
    evaluator.push_value(value)
}

fn observe_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    kernel: &mut StatisticalKernel,
    function: StatisticalFunction,
    value: WorkingValue<'a>,
    formula_error: &mut Option<ScalarError>,
    generated_error: &mut Option<ScalarError>,
) -> EvaluationResult<()> {
    match value {
        WorkingValue::Error(error) => {
            if function == StatisticalFunction::Count {
                // COUNT explicitly ignores formula errors.
            } else if function == StatisticalFunction::CountA {
                if let Err(error) = kernel.push_nonblank() {
                    remember(generated_error, error);
                }
            } else {
                remember(formula_error, error);
            }
        },
        WorkingValue::Number(value) => {
            let result = if function == StatisticalFunction::CountA {
                kernel.push_nonblank()
            } else {
                kernel.push_number(value)
            };
            if let Err(error) = result {
                remember(generated_error, error);
            }
        },
        WorkingValue::Logical(value) => {
            let result = if function == StatisticalFunction::CountA {
                kernel.push_nonblank()
            } else if function != StatisticalFunction::CountBlank {
                kernel.push_number(f64::from(value))
            } else {
                Ok(())
            };
            if let Err(error) = result {
                remember(generated_error, error);
            }
        },
        WorkingValue::Text(text) => {
            if function == StatisticalFunction::CountA {
                evaluator.charge_bytes(text.text.len())?;
                if let Err(error) = kernel.push_nonblank() {
                    remember(generated_error, error);
                }
            } else if function.is_a() {
                evaluator.charge_bytes(text.text.len())?;
                if let Err(error) = kernel.push_number(0.0) {
                    remember(generated_error, error);
                }
            } else {
                match super::to_number(WorkingValue::Text(text), evaluator)? {
                    Ok(value) => {
                        if let Err(error) = kernel.push_number(value) {
                            remember(generated_error, error);
                        }
                    },
                    Err(_error) if function == StatisticalFunction::Count => {
                        // COUNT ignores non-numeric direct values and their
                        // failed Number conversion just as it ignores Error
                        // values in a reference.
                    },
                    Err(error) => remember(generated_error, error),
                }
            }
        },
        WorkingValue::Complex(_) => {
            if function == StatisticalFunction::Count {
                // COUNT does not propagate non-numeric values.
            } else if function == StatisticalFunction::CountA {
                if let Err(error) = kernel.push_nonblank() {
                    remember(generated_error, error);
                }
            } else {
                remember(generated_error, ScalarError::Value);
            }
        },
    }
    Ok(())
}

fn remember(slot: &mut Option<ScalarError>, error: ScalarError) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn reverse_value_tail<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    count: usize,
) -> EvaluationResult<()> {
    evaluator.charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "statistical value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}
