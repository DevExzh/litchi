//! Scalar façade for the OpenFormula paired statistics.
//!
//! The value evaluator owns ForceArray shape validation, pairwise reference
//! traversal, and resolver accounting.  This module owns the function
//! identity and the resolver-free scalar bridge.  The fixed-width arithmetic
//! lives in [`kernel`], which is also shared by the value evaluator.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, UnsupportedKind,
    WorkingValue,
};

mod kernel;

pub(super) use kernel::{ExactPaired, ForecastFit, PairError, PairNeed};

/// The eight paired and simple-regression functions in the selected ODF
/// profile.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum PairedFunction {
    Correl,
    Covar,
    Pearson,
    Rsq,
    Slope,
    Intercept,
    Steyx,
    Forecast,
}

impl PairedFunction {
    /// Resolve one exact case-insensitive OpenFormula function name.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        Some(match name.len() {
            3 if name.eq_ignore_ascii_case("RSQ") => Self::Rsq,
            5 if name.eq_ignore_ascii_case("COVAR") => Self::Covar,
            5 if name.eq_ignore_ascii_case("SLOPE") => Self::Slope,
            5 if name.eq_ignore_ascii_case("STEYX") => Self::Steyx,
            6 if name.eq_ignore_ascii_case("CORREL") => Self::Correl,
            7 if name.eq_ignore_ascii_case("PEARSON") => Self::Pearson,
            8 if name.eq_ignore_ascii_case("FORECAST") => Self::Forecast,
            9 if name.eq_ignore_ascii_case("INTERCEPT") => Self::Intercept,
            _ => return None,
        })
    }

    /// Return whether an argument is a complete paired data argument.
    ///
    /// `FORECAST` has one scalar query at index zero and its Y/X data arrays
    /// at indices one and two.  Every other function has two data arguments.
    pub(super) const fn data_argument(self, index: usize) -> bool {
        match self {
            Self::Forecast => index != 0,
            _ => index < 2,
        }
    }

    /// Return the AST index for one of the two paired data arguments.
    ///
    /// `ordinal` is zero for the first data array and one for the second;
    /// `FORECAST` reserves AST index zero for its scalar query.
    pub(super) const fn data_argument_index(self, ordinal: usize) -> usize {
        match self {
            Self::Forecast => ordinal.saturating_add(1),
            _ => ordinal,
        }
    }

    /// Return the strict source arity required by this function.
    pub(super) const fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Forecast => count == 3,
            _ => count == 2,
        }
    }

    pub(super) const fn pair_need(self) -> PairNeed {
        match self {
            Self::Covar => PairNeed::Covariance,
            Self::Correl | Self::Pearson | Self::Rsq => PairNeed::Correlation,
            Self::Slope | Self::Intercept | Self::Forecast => PairNeed::Regression,
            Self::Steyx => PairNeed::Correlation,
        }
    }

    /// Alias retained for scalar callers that construct the kernel directly.
    const fn need(self) -> PairNeed {
        self.pair_need()
    }

    /// Return the data argument used as the independent X coordinate.
    pub(super) const fn x_index(self) -> usize {
        match self {
            Self::Correl | Self::Covar | Self::Pearson => 0,
            Self::Rsq | Self::Slope | Self::Intercept | Self::Steyx => 1,
            Self::Forecast => 2,
        }
    }

    /// Return the data argument used as the dependent Y coordinate.
    pub(super) const fn y_index(self) -> usize {
        match self {
            Self::Correl | Self::Covar | Self::Pearson => 1,
            Self::Rsq | Self::Slope | Self::Intercept | Self::Steyx => 0,
            Self::Forecast => 1,
        }
    }
}

/// Return whether `name` is one of the paired-statistics functions.
pub(super) fn is_paired_function(name: &str) -> bool {
    PairedFunction::from_name(name).is_some()
}

/// Keep the scalar stack mutation local to this bridge.  The source VM leaves
/// eager argument values in source order; reversing the borrowed tail lets us
/// pop without allocating a second `WorkingValue` vector.
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
                "paired value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

struct State {
    function: PairedFunction,
    kernel: ExactPaired,
    values: [Option<f64>; 3],
    query: Option<f64>,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
}

impl State {
    fn new(function: PairedFunction) -> Self {
        Self {
            function,
            kernel: ExactPaired::with_need(function.need()),
            values: [None; 3],
            query: None,
            formula_error: None,
            generated_error: None,
        }
    }

    fn remember(slot: &mut Option<ScalarError>, error: ScalarError) {
        if slot.is_none() {
            *slot = Some(error);
        }
    }

    fn formula(&mut self, error: ScalarError) {
        Self::remember(&mut self.formula_error, error);
    }

    fn generated(&mut self, error: ScalarError) {
        Self::remember(&mut self.generated_error, error);
    }

    fn finish_error(&self) -> Option<ScalarError> {
        self.formula_error.or(self.generated_error)
    }
}

/// Apply one paired function after eager scalar argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = PairedFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Function));
    };
    let count = node.child_count();
    if !function.valid_arity(count) {
        return evaluator.finish_invalid_arity(node);
    }

    reverse_value_tail(evaluator, count)?;
    let mut state = State::new(function);
    for index in 0..count {
        let value = evaluator.pop_value()?;
        if node.child(index).is_some_and(|child| child.is_missing()) {
            state.generated(ScalarError::Value);
            continue;
        }
        if function.data_argument(index) {
            observe_data(&mut state, index, value);
        } else {
            observe_query(evaluator, &mut state, value)?;
        }
    }

    // Observe every source argument before arithmetic. A retained error
    // already determines the result, so no numeric work is needed afterward.
    if state.finish_error().is_none()
        && let (Some(x), Some(y)) = (
            state.values[function.x_index()],
            state.values[function.y_index()],
        )
    {
        evaluator.charge_work(state.kernel.push_work_units())?;
        match state.kernel.push_pair(x, y) {
            Ok(()) => {},
            Err(error) => record_kernel_error(&mut state, error)?,
        }
    }

    let result = finish(evaluator, &mut state)?;
    evaluator.push_value(result)
}

fn observe_data(state: &mut State, index: usize, value: WorkingValue<'_>) {
    match value {
        WorkingValue::Number(value) if value.is_finite() => {
            if let Some(slot) = state.values.get_mut(index) {
                *slot = Some(value);
            } else {
                state.generated(ScalarError::Value);
            }
        },
        WorkingValue::Number(_) => state.generated(ScalarError::Number),
        // Text, Logical, and scalar Empty are omitted from a paired data
        // array.  The scalar VM has no Empty WorkingValue; the value VM keeps
        // it distinct and applies the same rule in its own bridge.
        WorkingValue::Text(_) | WorkingValue::Logical(_) => {},
        WorkingValue::Error(error) => state.formula(error),
        WorkingValue::Complex(_) => state.generated(ScalarError::Value),
    }
}

fn observe_query(
    evaluator: &mut Evaluator<'_, '_, '_>,
    state: &mut State,
    value: WorkingValue<'_>,
) -> EvaluationResult<()> {
    match value {
        WorkingValue::Number(value) if value.is_finite() => state.query = Some(value),
        WorkingValue::Number(_) => state.generated(ScalarError::Number),
        WorkingValue::Logical(value) => state.query = Some(f64::from(value)),
        WorkingValue::Text(text) => {
            evaluator.charge_bytes(text.text.len())?;
            match fast_float2::parse::<f64, _>(text.text.as_ref()).ok() {
                Some(value) if value.is_finite() => state.query = Some(value),
                Some(_) => state.generated(ScalarError::Number),
                None => state.generated(ScalarError::Value),
            }
        },
        WorkingValue::Error(error) => state.formula(error),
        WorkingValue::Complex(_) => state.generated(ScalarError::Value),
    }
    Ok(())
}

fn finish<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    state: &mut State,
) -> EvaluationResult<WorkingValue<'a>> {
    if let Some(error) = state.finish_error() {
        return Ok(WorkingValue::Error(error));
    }
    let work = if state.function == PairedFunction::Forecast {
        state
            .kernel
            .prepare_forecast_work_units()
            .checked_add(ForecastFit::query_work_units())
            .unwrap_or(u64::MAX)
    } else {
        state.kernel.finish_work_units()
    };
    evaluator.charge_work(work)?;
    let output = match state.function {
        PairedFunction::Covar => state.kernel.finish_covar(),
        PairedFunction::Correl | PairedFunction::Pearson => state.kernel.finish_correl(),
        PairedFunction::Rsq => state.kernel.finish_rsq(),
        PairedFunction::Slope => state.kernel.finish_slope(),
        PairedFunction::Intercept => state.kernel.finish_intercept(),
        PairedFunction::Steyx => state.kernel.finish_steyx(),
        PairedFunction::Forecast => match state.query {
            Some(query) => state.kernel.finish_forecast(query),
            None => Err(PairError::Empty),
        },
    };
    match output {
        Ok(value) => Ok(WorkingValue::Number(value)),
        Err(error) => {
            let error = map_function_pair_error(state.function, error);
            Ok(WorkingValue::Error(error))
        },
    }
}

fn record_kernel_error(state: &mut State, error: PairError) -> EvaluationResult<()> {
    state.generated(map_function_pair_error(state.function, error));
    Ok(())
}

fn map_function_pair_error(function: PairedFunction, error: PairError) -> ScalarError {
    match error {
        PairError::Empty => {
            if function == PairedFunction::Rsq {
                ScalarError::NotAvailable
            } else {
                ScalarError::Value
            }
        },
        PairError::MinimumCount => ScalarError::Value,
        PairError::ZeroVariance => ScalarError::DivisionByZero,
        PairError::Number => ScalarError::Number,
    }
}
