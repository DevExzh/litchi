//! Scalar kernels for the OpenFormula order and rank functions.
//!
//! This module owns the function identity, scalar argument handling, and the
//! numeric kernels shared by the resolver-backed value evaluator. The value
//! evaluator is responsible for walking references and arrays; once it has
//! admitted a finite `f64`, it can use [`OrderValues`] and the pure kernels
//! below without retaining a cell, text value, or resolver record.

use super::numerics::{NumericAggregate, NumericOperation, WideSum};
use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, WorkingValue,
    ensure_capacity,
};
use litchi_core::{Budget, ExecutionContext, Reservation};

/// The order and rank functions covered by this evaluator.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum OrderFunction {
    Median,
    Mode,
    Large,
    Small,
    Percentile,
    PercentRank,
    Quartile,
    Rank,
}

impl OrderFunction {
    /// Resolve one case-insensitive OpenFormula function name.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        let function = match name.len() {
            4 if name.eq_ignore_ascii_case("MODE") => Self::Mode,
            4 if name.eq_ignore_ascii_case("RANK") => Self::Rank,
            5 if name.eq_ignore_ascii_case("LARGE") => Self::Large,
            5 if name.eq_ignore_ascii_case("SMALL") => Self::Small,
            6 if name.eq_ignore_ascii_case("MEDIAN") => Self::Median,
            8 if name.eq_ignore_ascii_case("QUARTILE") => Self::Quartile,
            10 if name.eq_ignore_ascii_case("PERCENTILE") => Self::Percentile,
            11 if name.eq_ignore_ascii_case("PERCENTRANK") => Self::PercentRank,
            _ => return None,
        };
        Some(function)
    }

    /// Return whether a source argument is the sequence consumed by this
    /// function. Parameter positions intentionally remain position-sensitive
    /// for matrix demand and cache evaluation.
    pub(super) const fn data_argument(self, index: usize) -> bool {
        match self {
            Self::Median | Self::Mode => true,
            Self::Large | Self::Small | Self::Percentile | Self::PercentRank | Self::Quartile => {
                index == 0
            },
            Self::Rank => index == 1,
        }
    }

    /// Return the valid argument count for the scalar dispatcher.
    pub(super) fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Median | Self::Mode => count > 0,
            Self::Large | Self::Small | Self::Percentile | Self::Quartile => count == 2,
            Self::PercentRank | Self::Rank => (2..=3).contains(&count),
        }
    }
}

/// Return whether `name` is one of the order/rank functions.
pub(super) fn is_order_function(name: &str) -> bool {
    OrderFunction::from_name(name).is_some()
}

/// A fallibly reserved numeric buffer used by order reducers.
///
/// The vector is declared before its reservation so dropped values leave the
/// reservation only after the allocation has been released. `maximum` is
/// supplied by the caller: the scalar VM uses its AST stack bound, while the
/// value VM uses the admitted reference/array cell bound. In particular,
/// `max_stack_entries` is never imposed on a retained reference range merely
/// because this type is shared with scalar evaluation.
#[derive(Debug)]
pub(super) struct OrderValues {
    values: Vec<f64>,
    reservation: Option<Reservation>,
}

impl OrderValues {
    pub(super) fn new() -> Self {
        Self {
            values: Vec::new(),
            reservation: None,
        }
    }

    /// Append one finite value under the supplied live-storage bound.
    pub(super) fn push_with(
        &mut self,
        value: f64,
        maximum: usize,
        execution: &ExecutionContext,
        storage_budget: &Budget,
        scope: &'static str,
    ) -> EvaluationResult<()> {
        if !value.is_finite() {
            return Err(EvaluationFailure::InvalidExpression(
                "order reducer received a non-finite scalar",
            ));
        }
        ensure_capacity(
            &mut self.values,
            &mut self.reservation,
            1,
            maximum,
            execution,
            storage_budget,
            scope,
        )?;
        self.values.push(value);
        Ok(())
    }

    pub(super) fn as_slice(&self) -> &[f64] {
        &self.values
    }

    pub(super) fn as_mut_slice(&mut self) -> &mut [f64] {
        &mut self.values
    }

    pub(super) fn len(&self) -> usize {
        self.values.len()
    }
}

/// Sort finite numeric values in place with a bounded, fallible heapsort.
///
/// A standard library comparator cannot report cancellation or a work-limit
/// failure. The iterative heap implementation charges every comparison and
/// swap through `charge`, so callers retain a deterministic O(n log n) work
/// bound and cancellation checks throughout the sort.
pub(super) fn sort_values(
    values: &mut [f64],
    ascending: bool,
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<()> {
    if values.len() < 2 {
        return Ok(());
    }

    // Build a max heap for ascending output and a min heap for descending
    // output. Equal values compare equal numerically, including signed zero.
    let first = values.len() / 2;
    for start in (0..first).rev() {
        sift_down(values, start, values.len(), ascending, charge)?;
    }

    let mut end = values.len();
    while end > 1 {
        charge(1)?;
        values.swap(0, end - 1);
        end -= 1;
        sift_down(values, 0, end, ascending, charge)?;
    }
    Ok(())
}

fn sift_down(
    values: &mut [f64],
    mut root: usize,
    end: usize,
    ascending: bool,
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<()> {
    loop {
        let child = root
            .checked_mul(2)
            .and_then(|value| value.checked_add(1))
            .ok_or(EvaluationFailure::InvalidExpression(
                "order heap index overflow",
            ))?;
        if child >= end {
            return Ok(());
        }

        let mut selected = child;
        let right = child + 1;
        if right < end {
            charge(1)?;
            if precedes(values[right], values[child], ascending) {
                selected = right;
            }
        }
        charge(1)?;
        if !precedes(values[selected], values[root], ascending) {
            return Ok(());
        }
        charge(1)?;
        values.swap(root, selected);
        root = selected;
    }
}

/// Return whether `left` belongs above `right` in the heap used for the
/// requested final order. Numeric comparisons intentionally treat -0 and +0
/// as equal; no NaN can enter this module.
fn precedes(left: f64, right: f64, ascending: bool) -> bool {
    if ascending {
        left > right
    } else {
        left < right
    }
}

/// Compute MEDIAN from values already sorted in ascending numeric order.
pub(super) fn median_sorted(values: &[f64]) -> Result<f64, ScalarError> {
    if values.is_empty() {
        return Err(ScalarError::Value);
    }
    let middle = values.len() / 2;
    if values.len() % 2 == 1 {
        return Ok(canonical_zero(values[middle]));
    }

    let mut aggregate = NumericAggregate::new(NumericOperation::Average);
    aggregate.push_number(values[middle - 1])?;
    aggregate.push_number(values[middle])?;
    aggregate.result()?.ok_or(ScalarError::Value)
}

/// Compute MODE from values already sorted in ascending numeric order.
pub(super) fn mode_sorted(
    values: &[f64],
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if values.is_empty() {
        return Ok(Err(ScalarError::Value));
    }

    let mut best_count = 1usize;
    let mut best = values[0];
    let mut start = 0usize;
    while start < values.len() {
        let mut end = start + 1;
        while end < values.len() {
            charge(1)?;
            if values[end] != values[start] {
                break;
            }
            end += 1;
        }
        let count = end - start;
        // Ascending order means the first value wins a tie, including the
        // single numerical zero run containing both signs of zero.
        if count > best_count {
            best_count = count;
            best = values[start];
        }
        start = end;
    }

    if best_count < 2 {
        Ok(Err(ScalarError::Value))
    } else {
        Ok(Ok(canonical_zero(best)))
    }
}

/// Select one exact positive rank from ascending sorted values.
pub(super) fn select_sorted(values: &[f64], rank: f64, largest: bool) -> Result<f64, ScalarError> {
    let rank = exact_positive_integer(rank)?;
    if rank > values.len() as f64 {
        return Err(ScalarError::Value);
    }
    let rank = rank as usize;
    if rank == 0 || rank > values.len() {
        return Err(ScalarError::Value);
    }
    let index = if largest {
        values.len() - rank
    } else {
        rank - 1
    };
    values
        .get(index)
        .copied()
        .map(canonical_zero)
        .ok_or(ScalarError::Value)
}

/// Compute PERCENTILE from ascending sorted values.
pub(super) fn percentile_sorted(values: &[f64], x: f64) -> Result<f64, ScalarError> {
    if values.is_empty() || !x.is_finite() || !(0.0..=1.0).contains(&x) {
        return Err(ScalarError::Value);
    }

    // r = 1 + x * (n - 1), represented as zero-based rank here. Avoiding
    // the extra addition keeps endpoint classification exact for x=0 and 1.
    let zero_based = x * (values.len() - 1) as f64;
    let lower = zero_based.floor();
    let fraction = zero_based - lower;
    let index = usize::try_from(lower as u128).map_err(|_| ScalarError::Value)?;
    if fraction == 0.0 || index + 1 >= values.len() {
        return values
            .get(index)
            .copied()
            .map(canonical_zero)
            .ok_or(ScalarError::Value);
    }
    interpolate(values[index], values[index + 1], fraction)
}

/// Compute PERCENTRANK from ascending sorted values.
pub(super) fn percent_rank_sorted(
    values: &[f64],
    x: f64,
    significance: f64,
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if values.is_empty() || !x.is_finite() {
        return Ok(Err(ScalarError::Value));
    }
    let significance = match exact_positive_integer(significance) {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    if x < values[0] || x > values[values.len() - 1] {
        return Ok(Err(ScalarError::Value));
    }
    if values.len() == 1 {
        return Ok(round_percent_rank(1.0, significance));
    }

    // Lower-bound search gives duplicate X values their first occurrence
    // rank, as required by the OpenFormula profile.
    let rank = lower_bound(values, x, charge)?;
    let raw = if rank < values.len() && values[rank] == x {
        rank as f64 / (values.len() - 1) as f64
    } else {
        if rank == 0 || rank >= values.len() {
            return Ok(Err(ScalarError::Value));
        }
        let lower_value = values[rank - 1];
        let lower_rank = lower_bound(values, lower_value, charge)?;
        let fraction = match fraction_between(lower_value, x, values[rank]) {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        let fractional_rank = lower_rank as f64 + fraction;
        fractional_rank / (values.len() - 1) as f64
    };
    Ok(round_percent_rank(raw, significance))
}

/// Compute RANK without sorting. Equal values receive competition rank.
pub(super) fn rank_values(
    values: &[f64],
    value: f64,
    order: f64,
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    if !value.is_finite() || !order.is_finite() {
        return Ok(Err(ScalarError::Value));
    }
    let ascending = order != 0.0;
    let mut found = false;
    let mut preceding = 0usize;
    for candidate in values {
        charge(1)?;
        if *candidate == value {
            found = true;
        } else if (ascending && *candidate < value) || (!ascending && *candidate > value) {
            preceding = match preceding.checked_add(1) {
                Some(value) => value,
                None => return Ok(Err(ScalarError::Value)),
            };
        }
    }
    if !found {
        return Ok(Err(ScalarError::Value));
    }
    Ok(Ok((preceding + 1) as f64))
}

/// Validate a finite positive integer parameter.
pub(super) fn exact_positive_integer(value: f64) -> Result<f64, ScalarError> {
    if !value.is_finite() || value < 1.0 || value.fract() != 0.0 {
        return Err(ScalarError::Value);
    }
    Ok(value)
}

/// Validate the QUARTILE integer parameter and return its percentile.
pub(super) fn quartile_fraction(value: f64) -> Result<f64, ScalarError> {
    if !value.is_finite() || value.fract() != 0.0 || !(0.0..=4.0).contains(&value) {
        return Err(ScalarError::Value);
    }
    Ok(value / 4.0)
}

/// Overflow-safe interpolation of two finite ordered values.
pub(super) fn interpolate(lower: f64, upper: f64, fraction: f64) -> Result<f64, ScalarError> {
    if !lower.is_finite() || !upper.is_finite() || !fraction.is_finite() {
        return Err(ScalarError::Number);
    }
    if fraction == 0.0 {
        return Ok(canonical_zero(lower));
    }
    if fraction == 1.0 {
        return Ok(canonical_zero(upper));
    }
    let mut sum = WideSum::default();
    // lower + fraction * (upper - lower), expanded as three exact products
    // so an opposite-sign pair never overflows in an intermediate subtraction.
    sum.push_product(lower, 1.0)?;
    sum.push_product(upper, fraction)?;
    sum.push_product(lower, -fraction)?;
    sum.result()?.ok_or(ScalarError::Number)
}

/// Return the finite fraction of `value` between two ordered endpoints.
///
/// The direct subtraction path preserves adjacent-ULP precision. Scaling is
/// used only when an opposite-sign subtraction would overflow.
pub(super) fn fraction_between(lower: f64, value: f64, upper: f64) -> Result<f64, ScalarError> {
    if !(lower < value && value < upper) {
        return Err(ScalarError::Value);
    }
    let denominator = upper - lower;
    let numerator = value - lower;
    if denominator.is_finite() && numerator.is_finite() && denominator != 0.0 {
        let fraction = numerator / denominator;
        if fraction.is_finite() {
            return Ok(fraction);
        }
    }

    // Opposite-sign extrema can overflow the raw subtraction. Normalize only
    // in that case; retaining the raw path preserves adjacent-ULP precision
    // for large same-sign values.
    let scale = lower.abs().max(value.abs()).max(upper.abs());
    if scale == 0.0 || !scale.is_finite() {
        return Err(ScalarError::Number);
    }
    let lower_scaled = lower / scale;
    let value_scaled = value / scale;
    let upper_scaled = upper / scale;
    let denominator = upper_scaled - lower_scaled;
    let numerator = value_scaled - lower_scaled;
    if denominator == 0.0 {
        return Err(ScalarError::Number);
    }
    let fraction = numerator / denominator;
    if fraction.is_finite() {
        Ok(fraction)
    } else {
        Err(ScalarError::Number)
    }
}

fn lower_bound(
    values: &[f64],
    needle: f64,
    charge: &mut dyn FnMut(u64) -> EvaluationResult<()>,
) -> EvaluationResult<usize> {
    let mut low = 0usize;
    let mut high = values.len();
    while low < high {
        charge(1)?;
        let middle = low + (high - low) / 2;
        if values[middle] < needle {
            low = middle + 1;
        } else {
            high = middle;
        }
    }
    Ok(low)
}

fn round_percent_rank(value: f64, significance: f64) -> Result<f64, ScalarError> {
    // The rounding module owns the repository's decimal/tie profile. The
    // wrapper is intentionally small so both scalar and value evaluators use
    // exactly the same operation.
    super::rounding::round_nearest(value, significance)
}

fn canonical_zero(value: f64) -> f64 {
    if value == 0.0 { 0.0 } else { value }
}

#[derive(Debug)]
struct State {
    numbers: OrderValues,
    parameters: [Option<f64>; 3],
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
}

impl State {
    fn new() -> Self {
        Self {
            numbers: OrderValues::new(),
            parameters: [None; 3],
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

/// Apply one order/rank function after eager scalar argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let Some(function) = OrderFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::UnsupportedKind::Function,
        ));
    };
    let count = node.child_count();
    if !function.valid_arity(count) {
        return evaluator.finish_invalid_arity(node);
    }

    // Eager scheduling leaves the argument values in source order on the
    // evaluator stack. Reverse that borrowed tail and pop it one at a time;
    // retaining a second WorkingValue vector would double live storage for a
    // reducer whose data buffer is already the only required allocation.
    reverse_value_tail(evaluator, count)?;

    let mut state = State::new();
    for index in 0..count {
        let argument = evaluator.pop_value()?;
        // `visit_argument` turns a syntactic missing slot into the same
        // WorkingValue shape as a formula error. Keep its generated status
        // explicit so a later real formula error retains precedence.
        if node.child(index).is_some_and(|child| child.is_missing()) {
            state.generated(ScalarError::Value);
        } else if function.data_argument(index) {
            observe_data(evaluator, &mut state, argument)?;
        } else {
            observe_parameter(evaluator, &mut state, index, argument)?;
        }
    }

    let result = if let Some(error) = state.finish_error() {
        Ok(Err(error))
    } else {
        evaluate_result(
            evaluator,
            function,
            state.numbers.as_mut_slice(),
            &state.parameters,
        )
    };
    evaluator.push_value(match result? {
        Ok(value) => WorkingValue::Number(value),
        Err(error) => WorkingValue::Error(error),
    })
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
                "order value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

fn observe_data(
    evaluator: &mut Evaluator<'_, '_, '_>,
    state: &mut State,
    value: WorkingValue<'_>,
) -> EvaluationResult<()> {
    match value {
        WorkingValue::Number(value) => {
            if value.is_finite() {
                evaluator.charge_work(1)?;
                state.numbers.push_with(
                    value,
                    evaluator.limits.max_stack_entries,
                    evaluator.context.execution,
                    &evaluator.storage_budget,
                    "formula order numeric values",
                )?;
            } else {
                state.generated(ScalarError::Number);
            }
        },
        WorkingValue::Logical(value) => {
            evaluator.charge_work(1)?;
            state.numbers.push_with(
                f64::from(value),
                evaluator.limits.max_stack_entries,
                evaluator.context.execution,
                &evaluator.storage_budget,
                "formula order numeric values",
            )?;
        },
        WorkingValue::Text(text) => match super::to_number(WorkingValue::Text(text), evaluator)? {
            Ok(value) if value.is_finite() => {
                evaluator.charge_work(1)?;
                state.numbers.push_with(
                    value,
                    evaluator.limits.max_stack_entries,
                    evaluator.context.execution,
                    &evaluator.storage_budget,
                    "formula order numeric values",
                )?;
            },
            Ok(_) => state.generated(ScalarError::Number),
            Err(error) => state.generated(error),
        },
        WorkingValue::Error(error) => state.formula(error),
        WorkingValue::Complex(_) => state.generated(ScalarError::Value),
    }
    Ok(())
}

fn observe_parameter(
    evaluator: &mut Evaluator<'_, '_, '_>,
    state: &mut State,
    index: usize,
    value: WorkingValue<'_>,
) -> EvaluationResult<()> {
    if let WorkingValue::Error(error) = value {
        state.formula(error);
        return Ok(());
    }
    // Parameter conversion errors are generated errors. A formula Error was
    // handled above, preserving the distinction needed for precedence.
    match super::to_number(value, evaluator)? {
        Ok(value) if value.is_finite() => {
            if let Some(slot) = state.parameters.get_mut(index) {
                *slot = Some(value);
            } else {
                state.generated(ScalarError::Value);
            }
        },
        Ok(_) => state.generated(ScalarError::Number),
        Err(error) => state.generated(error),
    }
    Ok(())
}

fn evaluate_result(
    evaluator: &mut Evaluator<'_, '_, '_>,
    function: OrderFunction,
    values: &mut [f64],
    parameters: &[Option<f64>; 3],
) -> EvaluationResult<Result<f64, ScalarError>> {
    // Scalar evaluation has no reference/array descriptor, so every sort is
    // bounded by the already-admitted scalar argument count.
    let mut charge = |amount| evaluator.charge_work(amount);
    match function {
        OrderFunction::Median => {
            sort_values(values, true, &mut charge)?;
            Ok(median_sorted(values))
        },
        OrderFunction::Mode => {
            sort_values(values, true, &mut charge)?;
            mode_sorted(values, &mut charge)
        },
        OrderFunction::Large | OrderFunction::Small => {
            let rank = match scalar_parameter(parameters, 1) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            sort_values(values, true, &mut charge)?;
            Ok(select_sorted(
                values,
                rank,
                function == OrderFunction::Large,
            ))
        },
        OrderFunction::Percentile => {
            let x = match scalar_parameter(parameters, 1) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            sort_values(values, true, &mut charge)?;
            Ok(percentile_sorted(values, x))
        },
        OrderFunction::PercentRank => {
            let x = match scalar_parameter(parameters, 1) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            let significance = parameters[2].unwrap_or(3.0);
            sort_values(values, true, &mut charge)?;
            percent_rank_sorted(values, x, significance, &mut charge)
        },
        OrderFunction::Quartile => {
            let quart = match scalar_parameter(parameters, 1) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            let x = match quartile_fraction(quart) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            sort_values(values, true, &mut charge)?;
            Ok(percentile_sorted(values, x))
        },
        OrderFunction::Rank => {
            let value = match scalar_parameter(parameters, 0) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            };
            let order = parameters[2].unwrap_or(0.0);
            rank_values(values, value, order, &mut charge)
        },
    }
}

fn scalar_parameter(parameters: &[Option<f64>; 3], index: usize) -> Result<f64, ScalarError> {
    parameters
        .get(index)
        .and_then(|value| *value)
        .ok_or(ScalarError::Value)
}
