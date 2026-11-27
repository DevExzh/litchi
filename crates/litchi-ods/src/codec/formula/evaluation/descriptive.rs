//! Scalar kernels for the OpenFormula descriptive statistics.
//!
//! The resolver-backed value evaluator owns sequence admission and descriptor
//! replay.  This module owns the finite-number state used after admission and
//! the resolver-free scalar bridge.  Centered reducers retain exact bounded
//! dyadic raw moments, so forming one rounded absolute mean cannot lose the
//! half-ULP center of adjacent large values.

use super::numerics::{binary_parts, scale_binary};
use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, UnsupportedKind,
    WorkingValue,
};
use litchi_core::{Reservation, Resource};

pub(super) mod harmonic;
mod moments;

// The resolver-backed value profile is a sibling module.  Re-export the
// harmonic phase API at the evaluation crate boundary so that profile can
// share the exact replay without exposing the implementation submodule.
pub(super) use harmonic::{
    ExactHarmonicAccumulator, HarmonicFailure, HarmonicFirstPass, HarmonicKernel, recommended_limbs,
};

/// The descriptive functions covered by this evaluator.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum DescriptiveFunction {
    AveDev,
    DevSq,
    Geomean,
    Harmean,
    Kurt,
    Skew,
    SkewP,
}

impl DescriptiveFunction {
    /// Resolve one case-insensitive OpenFormula function name.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        Some(match name.len() {
            6 if name.eq_ignore_ascii_case("AVEDEV") => Self::AveDev,
            5 if name.eq_ignore_ascii_case("DEVSQ") => Self::DevSq,
            4 if name.eq_ignore_ascii_case("KURT") => Self::Kurt,
            4 if name.eq_ignore_ascii_case("SKEW") => Self::Skew,
            5 if name.eq_ignore_ascii_case("SKEWP") => Self::SkewP,
            7 if name.eq_ignore_ascii_case("GEOMEAN") => Self::Geomean,
            7 if name.eq_ignore_ascii_case("HARMEAN") => Self::Harmean,
            _ => return None,
        })
    }

    /// Whether this function admits an explicit ordered ReferenceList.
    pub(super) const fn accepts_reference_list(self) -> bool {
        !matches!(self, Self::DevSq | Self::SkewP)
    }

    /// Whether a second descriptor pass is required after the mean pass.
    pub(super) const fn needs_replay(self) -> bool {
        matches!(self, Self::AveDev)
    }
}

/// Return whether `name` is one of the descriptive reducers.
pub(super) fn is_descriptive_function(name: &str) -> bool {
    DescriptiveFunction::from_name(name).is_some()
}

/// Marker retained between the first and AVEDEV replay passes.
///
/// The exact mean and replay state live in [`moments::ExactMoments`] and its
/// fixed-size [`moments::AveDevReplay`].  Keeping this marker copyable lets
/// the scalar and resolver-backed bridges share the same pass API without
/// copying the bounded arithmetic state.
#[derive(Clone, Copy, Debug)]
pub(super) struct MeanPlan;

/// Shared fixed state for geometric and exact descriptive reducers.
///
/// `HARMEAN` has a separate arithmetic state in `harmonic.rs`; keeping it out
/// of this struct avoids growing every other reducer by its reciprocal profile.
pub(super) struct DescriptiveKernel {
    function: DescriptiveFunction,
    moments: moments::ExactMoments,
    geometric: GeometricAccumulator,
    ave_dev_replay: Option<moments::AveDevReplay>,
}

impl DescriptiveKernel {
    /// Create a fixed state for one descriptive reducer.
    pub(super) fn new(function: DescriptiveFunction) -> Self {
        let order = match function {
            DescriptiveFunction::AveDev | DescriptiveFunction::Geomean => {
                moments::MomentOrder::First
            },
            DescriptiveFunction::DevSq => moments::MomentOrder::Second,
            DescriptiveFunction::Kurt => moments::MomentOrder::Fourth,
            DescriptiveFunction::Skew | DescriptiveFunction::SkewP => moments::MomentOrder::Third,
            DescriptiveFunction::Harmean => moments::MomentOrder::First,
        };
        Self {
            function,
            moments: moments::ExactMoments::with_order(order),
            geometric: GeometricAccumulator::default(),
            ave_dev_replay: None,
        }
    }

    /// Add one admitted finite Number during the first pass.
    pub(super) fn push_first(&mut self, value: f64) -> Result<(), ScalarError> {
        match self.function {
            DescriptiveFunction::Geomean => self.geometric.push(value),
            DescriptiveFunction::AveDev
            | DescriptiveFunction::DevSq
            | DescriptiveFunction::Kurt
            | DescriptiveFunction::Skew
            | DescriptiveFunction::SkewP => self
                .moments
                .push(value)
                .map_err(|error| map_moment_error(self.function, error)),
            DescriptiveFunction::Harmean => Err(ScalarError::Value),
        }
    }

    /// Prepare the second pass.  AVEDEV needs the exact mean to form absolute
    /// deviations; the other four centered reducers finish from their exact
    /// raw moments after one scan.
    pub(super) fn prepare_replay(&mut self) -> Result<Option<MeanPlan>, ScalarError> {
        if !self.function.needs_replay() {
            return Ok(None);
        }
        let plan = self
            .moments
            .ave_dev_plan()
            .map_err(|error| map_moment_error(self.function, error))?;
        self.ave_dev_replay = Some(plan.replay());
        Ok(Some(MeanPlan))
    }

    /// Add one value during the AVEDEV replay pass.
    pub(super) fn push_replay(&mut self, _plan: MeanPlan, value: f64) -> Result<(), ScalarError> {
        if !self.function.needs_replay() {
            return Ok(());
        }
        let replay = self.ave_dev_replay.as_mut().ok_or(ScalarError::Value)?;
        replay
            .push(value)
            .map_err(|error| map_moment_error(self.function, error))
    }

    /// Finish the reducer after all required passes.
    pub(super) fn finish(&self) -> Result<f64, ScalarError> {
        match self.function {
            DescriptiveFunction::AveDev => self
                .ave_dev_replay
                .as_ref()
                .ok_or(ScalarError::DivisionByZero)?
                .finish()
                .map_err(|error| map_moment_error(self.function, error)),
            DescriptiveFunction::DevSq => self
                .moments
                .finish_devsq()
                .map_err(|error| map_moment_error(self.function, error)),
            DescriptiveFunction::Geomean => self.geometric.finish(),
            DescriptiveFunction::Harmean => Err(ScalarError::DivisionByZero),
            DescriptiveFunction::Kurt => self
                .moments
                .finish_kurt()
                .map_err(|error| map_moment_error(self.function, error)),
            DescriptiveFunction::Skew => self
                .moments
                .finish_skew(true)
                .map_err(|error| map_moment_error(self.function, error)),
            DescriptiveFunction::SkewP => self
                .moments
                .finish_skew(false)
                .map_err(|error| map_moment_error(self.function, error)),
        }
    }
}

fn map_moment_error(function: DescriptiveFunction, error: moments::MomentError) -> ScalarError {
    match error {
        moments::MomentError::Empty => match function {
            DescriptiveFunction::AveDev | DescriptiveFunction::DevSq => ScalarError::DivisionByZero,
            DescriptiveFunction::Geomean | DescriptiveFunction::Harmean => {
                ScalarError::DivisionByZero
            },
            DescriptiveFunction::Kurt | DescriptiveFunction::Skew | DescriptiveFunction::SkewP => {
                ScalarError::Value
            },
        },
        moments::MomentError::MinimumCount | moments::MomentError::ZeroVariance => {
            ScalarError::Value
        },
        moments::MomentError::Number => ScalarError::Number,
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct GeometricAccumulator {
    count: u64,
    negative_count: u64,
    zero: bool,
    first_absolute: f64,
    all_absolute_equal: bool,
    reference_mantissa: f64,
    reference_exponent: i32,
    ratio_mantissa: f64,
    ratio_exponent: i128,
}

impl GeometricAccumulator {
    fn push(&mut self, value: f64) -> Result<(), ScalarError> {
        if !value.is_finite() {
            return Err(ScalarError::Number);
        }
        self.count = self.count.checked_add(1).ok_or(ScalarError::Number)?;
        if value == 0.0 {
            self.zero = true;
            return Ok(());
        }
        // The exact product is already zero.  Keep validating the finite
        // boundary above and the count, but no later magnitude can affect
        // the canonical zero result.
        if self.zero {
            return Ok(());
        }
        if self.count == 1 {
            self.first_absolute = value.abs();
            self.all_absolute_equal = true;
            let (mantissa, exponent) = binary_parts(self.first_absolute);
            self.reference_mantissa = mantissa;
            self.reference_exponent = exponent;
            self.ratio_mantissa = 1.0;
            self.ratio_exponent = 0;
        } else if value.abs() != self.first_absolute {
            self.all_absolute_equal = false;
        }
        if value.is_sign_negative() {
            self.negative_count = self
                .negative_count
                .checked_add(1)
                .ok_or(ScalarError::Number)?;
        }
        // Keep the product relative to the first magnitude.  Multiplying
        // normalized mantissas avoids both the loss from averaging a large
        // absolute exponent and the `exp2(log2(x))` round trip error for a
        // cluster of adjacent values.  The exponent is kept separately, so
        // the product itself never overflows.
        if self.reference_mantissa == 0.0 {
            let (mantissa, exponent) = binary_parts(value.abs());
            self.first_absolute = value.abs();
            self.all_absolute_equal = true;
            self.reference_mantissa = mantissa;
            self.reference_exponent = exponent;
            return Ok(());
        }
        let (mantissa, exponent) = binary_parts(value.abs());
        let mut ratio_exponent = i64::from(exponent)
            .checked_sub(i64::from(self.reference_exponent))
            .ok_or(ScalarError::Number)?;
        let ratio = mantissa / self.reference_mantissa;
        let mut ratio_mantissa = ratio;
        normalize_mantissa(&mut ratio_mantissa, &mut ratio_exponent)?;
        self.ratio_mantissa *= ratio_mantissa;
        self.ratio_exponent = self
            .ratio_exponent
            .checked_add(i128::from(ratio_exponent))
            .ok_or(ScalarError::Number)?;
        if !self.ratio_mantissa.is_finite() || self.ratio_mantissa <= 0.0 {
            return Err(ScalarError::Number);
        }
        while self.ratio_mantissa >= 2.0 {
            self.ratio_mantissa *= 0.5;
            self.ratio_exponent = self
                .ratio_exponent
                .checked_add(1)
                .ok_or(ScalarError::Number)?;
        }
        while self.ratio_mantissa < 1.0 {
            self.ratio_mantissa *= 2.0;
            self.ratio_exponent = self
                .ratio_exponent
                .checked_sub(1)
                .ok_or(ScalarError::Number)?;
        }
        Ok(())
    }

    fn finish(&self) -> Result<f64, ScalarError> {
        if self.count == 0 {
            return Err(ScalarError::DivisionByZero);
        }
        if self.zero {
            return Ok(0.0);
        }
        if self.negative_count % 2 == 1 && self.count % 2 == 0 {
            return Err(ScalarError::Number);
        }
        if self.all_absolute_equal {
            return Ok(if self.negative_count % 2 == 1 {
                -self.first_absolute
            } else {
                self.first_absolute
            });
        }
        let count = i128::from(self.count);
        let root_exponent = self.ratio_exponent.div_euclid(count);
        let remainder = self.ratio_exponent.rem_euclid(count);
        let logarithm = (remainder as f64 + self.ratio_mantissa.log2()) / self.count as f64;
        if !logarithm.is_finite() {
            return Err(ScalarError::Number);
        }
        let mut ratio_root = logarithm.exp2();
        let mut root_exponent = i64::try_from(root_exponent).map_err(|_| ScalarError::Number)?;
        normalize_mantissa(&mut ratio_root, &mut root_exponent)?;

        let mut mantissa = self.reference_mantissa * ratio_root;
        let mut exponent = i64::from(self.reference_exponent)
            .checked_add(root_exponent)
            .ok_or(ScalarError::Number)?;
        normalize_mantissa(&mut mantissa, &mut exponent)?;
        let magnitude = scale_binary(mantissa, exponent);
        if !magnitude.is_finite() && magnitude != 0.0 {
            return Err(ScalarError::Number);
        }
        Ok(if self.negative_count % 2 == 1 {
            -magnitude
        } else {
            magnitude
        })
    }
}

fn normalize_mantissa(mantissa: &mut f64, exponent: &mut i64) -> Result<(), ScalarError> {
    if !mantissa.is_finite() || *mantissa <= 0.0 {
        return Err(ScalarError::Number);
    }
    while *mantissa >= 2.0 {
        *mantissa *= 0.5;
        *exponent = exponent.checked_add(1).ok_or(ScalarError::Number)?;
    }
    while *mantissa < 1.0 {
        *mantissa *= 2.0;
        *exponent = exponent.checked_sub(1).ok_or(ScalarError::Number)?;
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
enum Observation {
    Number(f64),
    FormulaError(ScalarError),
    GeneratedError(ScalarError),
}

fn observe_ref(value: &WorkingValue<'_>) -> (Observation, usize) {
    match value {
        WorkingValue::Number(value) if value.is_finite() => (Observation::Number(*value), 0),
        WorkingValue::Number(_) => (Observation::GeneratedError(ScalarError::Number), 0),
        WorkingValue::Logical(value) => (Observation::Number(f64::from(*value)), 0),
        WorkingValue::Text(text) => {
            let bytes = text.text.len();
            let parsed = fast_float2::parse::<f64, _>(text.text.as_ref()).ok();
            match parsed.filter(|value| value.is_finite()) {
                Some(value) => (Observation::Number(value), bytes),
                None => (Observation::GeneratedError(ScalarError::Value), bytes),
            }
        },
        WorkingValue::Error(error) => (Observation::FormulaError(*error), 0),
        WorkingValue::Complex(_) => (Observation::GeneratedError(ScalarError::Value), 0),
    }
}

fn observe_owned(value: WorkingValue<'_>) -> (Observation, usize) {
    match value {
        WorkingValue::Number(value) if value.is_finite() => (Observation::Number(value), 0),
        WorkingValue::Number(_) => (Observation::GeneratedError(ScalarError::Number), 0),
        WorkingValue::Logical(value) => (Observation::Number(f64::from(value)), 0),
        WorkingValue::Text(text) => {
            let bytes = text.text.len();
            let parsed = fast_float2::parse::<f64, _>(text.text.as_ref()).ok();
            match parsed.filter(|value| value.is_finite()) {
                Some(value) => (Observation::Number(value), bytes),
                None => (Observation::GeneratedError(ScalarError::Value), bytes),
            }
        },
        WorkingValue::Error(error) => (Observation::FormulaError(error), 0),
        WorkingValue::Complex(_) => (Observation::GeneratedError(ScalarError::Value), 0),
    }
}

/// Apply one descriptive reducer after eager scalar argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = DescriptiveFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Function));
    };
    if function == DescriptiveFunction::Harmean {
        return apply_harmonic(evaluator, node);
    }

    let count = node.child_count();
    if count == 0 {
        return evaluator.finish_invalid_arity(node);
    }
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "descriptive value stack underflow",
            ))?;

    let mut kernel = DescriptiveKernel::new(function);
    let mut formula_error = None;
    let mut generated_error = None;
    for index in start..evaluator.values.len() {
        evaluator.charge_work(1)?;
        let (observation, bytes) = observe_ref(&evaluator.values[index]);
        if bytes != 0 {
            evaluator.charge_bytes(bytes)?;
        }
        observe_first(
            &mut kernel,
            observation,
            &mut formula_error,
            &mut generated_error,
        );
    }

    let replay = if formula_error.is_none() && generated_error.is_none() {
        match kernel.prepare_replay() {
            Ok(replay) => replay,
            Err(error) => {
                remember(&mut generated_error, error);
                None
            },
        }
    } else {
        None
    };

    reverse_value_tail(evaluator, count)?;
    if let Some(plan) = replay {
        for _ in 0..count {
            evaluator.charge_work(1)?;
            let (observation, bytes) = observe_owned(evaluator.pop_value()?);
            if bytes != 0 {
                evaluator.charge_bytes(bytes)?;
            }
            if let Observation::Number(value) = observation {
                if let Err(error) = kernel.push_replay(plan, value) {
                    remember(&mut generated_error, error);
                }
            }
        }
    } else {
        for _ in 0..count {
            evaluator.charge_work(1)?;
            let _ = evaluator.pop_value()?;
        }
    }

    let result = if let Some(error) = formula_error.or(generated_error) {
        WorkingValue::Error(error)
    } else {
        evaluator.charge_work(1)?;
        match kernel.finish() {
            Ok(value) => WorkingValue::Number(value),
            Err(error) => WorkingValue::Error(error),
        }
    };
    evaluator.push_value(result)
}

fn apply_harmonic<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    let count = node.child_count();
    if count == 0 {
        return evaluator.finish_invalid_arity(node);
    }
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "harmonic value stack underflow",
            ))?;
    let mut fast = HarmonicKernel::new();
    let mut formula_error = None;
    let mut generated_error = None;
    for index in start..evaluator.values.len() {
        evaluator.charge_work(1)?;
        let (observation, bytes) = observe_ref(&evaluator.values[index]);
        if bytes != 0 {
            evaluator.charge_bytes(bytes)?;
        }
        match observation {
            Observation::Number(value) => {
                if let Err(error) = fast.push_number(value) {
                    if let Some(error) = harmonic_formula_error(error) {
                        remember(&mut generated_error, error);
                    }
                }
            },
            Observation::FormulaError(error) => remember(&mut formula_error, error),
            Observation::GeneratedError(error) => remember(&mut generated_error, error),
        }
    }

    let first_pass = if formula_error.is_none() && generated_error.is_none() {
        match fast.finish() {
            Ok(outcome) => Some(outcome),
            Err(error) => {
                if let Some(error) = harmonic_formula_error(error) {
                    remember(&mut generated_error, error);
                }
                None
            },
        }
    } else {
        None
    };

    let mut resource_failure = None;
    let mut exact_reservation = None;
    let mut exact = if matches!(first_pass, Some(HarmonicFirstPass::Replay)) {
        match recommended_limbs(fast.count()) {
            Ok(max_limbs) => match ExactHarmonicAccumulator::new(max_limbs) {
                Ok(accumulator) => Some(accumulator),
                Err(error) => {
                    remember_harmonic_failure(
                        error,
                        &mut generated_error,
                        &mut resource_failure,
                        evaluator.limits.max_storage_bytes,
                    );
                    None
                },
            },
            Err(error) => {
                remember_harmonic_failure(
                    error,
                    &mut generated_error,
                    &mut resource_failure,
                    evaluator.limits.max_storage_bytes,
                );
                None
            },
        }
    } else {
        None
    };

    reverse_value_tail(evaluator, count)?;
    let mut replay = exact.is_some();
    for _ in 0..count {
        evaluator.charge_work(1)?;
        if !replay {
            let _ = evaluator.pop_value()?;
            continue;
        }
        let (observation, bytes) = observe_owned(evaluator.pop_value()?);
        if bytes != 0 {
            evaluator.charge_bytes(bytes)?;
        }
        let mut exact_error = None;
        let mut storage_error = None;
        if let Some(accumulator) = exact.as_ref()
            && let Observation::Number(_) = observation
        {
            match accumulator.push_work_units() {
                Ok(work) => evaluator.charge_work(work)?,
                Err(error) => exact_error = Some(error),
            }
        }
        if exact_error.is_none()
            && let Some(accumulator) = exact.as_ref()
            && let Observation::Number(value) = observation
        {
            match accumulator.push_storage_bytes(value) {
                Ok(required_bytes) => {
                    if let Err(error) =
                        reserve_harmonic_storage(evaluator, &mut exact_reservation, required_bytes)
                    {
                        storage_error = Some(error);
                    }
                },
                Err(error) => exact_error = Some(error),
            }
        }
        if let Some(accumulator) = exact.as_mut()
            && exact_error.is_none()
            && storage_error.is_none()
            && let Observation::Number(value) = observation
        {
            if let Err(error) = accumulator.push_number(value) {
                exact_error = Some(error);
            }
        }
        if let Some(error) = exact_error {
            let typed_failure = matches!(
                error,
                HarmonicFailure::Resource | HarmonicFailure::Allocation { .. }
            );
            remember_harmonic_failure(
                error,
                &mut generated_error,
                &mut resource_failure,
                evaluator.limits.max_storage_bytes,
            );
            if typed_failure || generated_error.is_some() {
                exact = None;
                replay = false;
            }
        }
    }

    if let Some(error) = resource_failure {
        return Err(error);
    }

    let result = if let Some(error) = formula_error.or(generated_error) {
        WorkingValue::Error(error)
    } else {
        evaluator.charge_work(1)?;
        match first_pass {
            Some(HarmonicFirstPass::Value(value)) => WorkingValue::Number(value),
            Some(HarmonicFirstPass::Replay) => match exact {
                Some(mut exact) => {
                    if exact.count() != fast.count() {
                        return Err(EvaluationFailure::InvalidExpression(
                            "harmonic replay count diverged",
                        ));
                    }
                    let finish_storage = match exact.finish_storage_bytes() {
                        Ok(bytes) => bytes,
                        Err(error) => {
                            return Err(harmonic_evaluation_failure(
                                error,
                                evaluator.limits.max_storage_bytes,
                            ));
                        },
                    };
                    reserve_harmonic_storage(evaluator, &mut exact_reservation, finish_storage)?;
                    let finish_work = match exact.finish_work_units() {
                        Ok(work) => work,
                        Err(error) => {
                            return Err(harmonic_evaluation_failure(
                                error,
                                evaluator.limits.max_storage_bytes,
                            ));
                        },
                    };
                    evaluator.charge_work(finish_work)?;
                    match exact.finish() {
                        Ok(value) => WorkingValue::Number(value),
                        Err(
                            error
                            @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. }),
                        ) => {
                            return Err(harmonic_evaluation_failure(
                                error,
                                evaluator.limits.max_storage_bytes,
                            ));
                        },
                        Err(error) => WorkingValue::Error(harmonic_formula_error(error).ok_or(
                            EvaluationFailure::InvalidExpression(
                                "harmonic formula failure was not surfaced",
                            ),
                        )?),
                    }
                },
                None => WorkingValue::Error(ScalarError::Number),
            },
            None => WorkingValue::Error(ScalarError::DivisionByZero),
        }
    };
    evaluator.push_value(result)
}

fn harmonic_formula_error(error: HarmonicFailure) -> Option<ScalarError> {
    match error {
        HarmonicFailure::DivisionByZero => Some(ScalarError::DivisionByZero),
        HarmonicFailure::Number => Some(ScalarError::Number),
        HarmonicFailure::Resource | HarmonicFailure::Allocation { .. } => None,
    }
}

fn harmonic_evaluation_failure(error: HarmonicFailure, storage_limit: usize) -> EvaluationFailure {
    match error {
        HarmonicFailure::Resource => super::local_limit(
            Resource::Memory,
            u64::try_from(storage_limit)
                .unwrap_or(u64::MAX)
                .saturating_add(1),
            u64::try_from(storage_limit).unwrap_or(u64::MAX),
        ),
        HarmonicFailure::Allocation { source } => EvaluationFailure::Allocation {
            resource: "formula harmonic exact replay",
            source,
        },
        HarmonicFailure::DivisionByZero | HarmonicFailure::Number => {
            EvaluationFailure::InvalidExpression("harmonic formula failure was not typed")
        },
    }
}

fn remember_harmonic_failure(
    error: HarmonicFailure,
    generated_error: &mut Option<ScalarError>,
    resource_failure: &mut Option<EvaluationFailure>,
    storage_limit: usize,
) {
    match error {
        HarmonicFailure::Resource => {
            if resource_failure.is_none() {
                *resource_failure = Some(super::local_limit(
                    Resource::Memory,
                    u64::try_from(storage_limit)
                        .unwrap_or(u64::MAX)
                        .saturating_add(1),
                    u64::try_from(storage_limit).unwrap_or(u64::MAX),
                ));
            }
        },
        HarmonicFailure::Allocation { source } => {
            if resource_failure.is_none() {
                *resource_failure = Some(EvaluationFailure::Allocation {
                    resource: "formula harmonic exact replay",
                    source,
                });
            }
        },
        error => {
            if let Some(error) = harmonic_formula_error(error) {
                remember(generated_error, error);
            }
        },
    }
}

fn reserve_harmonic_storage(
    evaluator: &mut Evaluator<'_, '_, '_>,
    reservation: &mut Option<Reservation>,
    required: usize,
) -> EvaluationResult<()> {
    let current = reservation.as_ref().map_or(0, |reservation| {
        usize::try_from(reservation.amount()).unwrap_or(usize::MAX)
    });
    if required <= current {
        return Ok(());
    }
    let additional = required - current;
    let next = evaluator.reserve_storage(additional, "formula harmonic exact replay")?;
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

fn observe_first(
    kernel: &mut DescriptiveKernel,
    observation: Observation,
    formula_error: &mut Option<ScalarError>,
    generated_error: &mut Option<ScalarError>,
) {
    match observation {
        Observation::Number(value) => {
            if let Err(error) = kernel.push_first(value) {
                remember(generated_error, error);
            }
        },
        Observation::FormulaError(error) => remember(formula_error, error),
        Observation::GeneratedError(error) => remember(generated_error, error),
    }
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
                "descriptive value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}
