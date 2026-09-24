//! Resolver-aware numeric sequence and matrix aggregates.
//!
//! The scalar evaluator owns the scalar kernels.  This module owns the value
//! profile's two sequence shapes: `SUM`, `PRODUCT`, `SUMSQ`, `GCD`, `LCM`, and
//! `MULTINOMIAL` consume ordered number sequences, while `SUMPRODUCT` and the three `SUMX*`
//! functions consume equal-sized forced arrays.  References are traversed in
//! place, so a large range does not first become a second owned array.

use super::super::discrete::{DiscreteFold, DiscreteFunction, sequence_function};
use super::super::numerics::{
    NumericAggregate, NumericOperation, ProductSumError, ProductTerm, ScaledProductSum, WideSum,
};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, Resource, RuntimeAreaSet, RuntimeElement,
    RuntimeValue, ScalarError, Shape, ValueEvaluator, WorkingValue,
};

/// Numeric sequence and matrix functions covered by this value bridge.
pub(super) fn is_aggregate_function(name: &str) -> bool {
    name.eq_ignore_ascii_case("SUM")
        || name.eq_ignore_ascii_case("PRODUCT")
        || name.eq_ignore_ascii_case("SUMSQ")
        || sequence_function(name).is_some()
        || name.eq_ignore_ascii_case("SUMPRODUCT")
        || name.eq_ignore_ascii_case("SUMX2MY2")
        || name.eq_ignore_ascii_case("SUMX2PY2")
        || name.eq_ignore_ascii_case("SUMXMY2")
}

/// Return whether the function's parameters use the normative `ForceArray`
/// signature.  Force-array parameters must retain a complete array even when
/// the surrounding value evaluation is in scalar mode.
pub(super) fn is_matrix_aggregate_function(name: &str) -> bool {
    name.eq_ignore_ascii_case("SUMPRODUCT")
        || name.eq_ignore_ascii_case("SUMX2MY2")
        || name.eq_ignore_ascii_case("SUMX2PY2")
        || name.eq_ignore_ascii_case("SUMXMY2")
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Function {
    Sum,
    Product,
    SumSquares,
    Gcd,
    Lcm,
    Multinomial,
    SumProduct,
    SumX2My2,
    SumX2Py2,
    SumXMy2,
}

impl Function {
    fn from_name(name: &str) -> Option<Self> {
        Some(if name.eq_ignore_ascii_case("SUM") {
            Self::Sum
        } else if name.eq_ignore_ascii_case("PRODUCT") {
            Self::Product
        } else if name.eq_ignore_ascii_case("SUMSQ") {
            Self::SumSquares
        } else if name.eq_ignore_ascii_case("GCD") {
            Self::Gcd
        } else if name.eq_ignore_ascii_case("LCM") {
            Self::Lcm
        } else if name.eq_ignore_ascii_case("MULTINOMIAL") {
            Self::Multinomial
        } else if name.eq_ignore_ascii_case("SUMPRODUCT") {
            Self::SumProduct
        } else if name.eq_ignore_ascii_case("SUMX2MY2") {
            Self::SumX2My2
        } else if name.eq_ignore_ascii_case("SUMX2PY2") {
            Self::SumX2Py2
        } else if name.eq_ignore_ascii_case("SUMXMY2") {
            Self::SumXMy2
        } else {
            return None;
        })
    }

    const fn is_sequence(self) -> bool {
        matches!(
            self,
            Self::Sum
                | Self::Product
                | Self::SumSquares
                | Self::Gcd
                | Self::Lcm
                | Self::Multinomial
        )
    }

    const fn is_pair(self) -> bool {
        matches!(self, Self::SumX2My2 | Self::SumX2Py2 | Self::SumXMy2)
    }
}

/// Apply one value-profile aggregate after its arguments have been visited.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
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
    if arguments
        .iter()
        .any(|value| matches!(value, RuntimeValue::SourceReference))
    {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        ));
    }
    if function.is_pair() && arguments.len() != 2 {
        if let Some(error) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }
    if matches!(
        function,
        Function::Product | Function::Gcd | Function::Lcm | Function::Multinomial
    ) && arguments.is_empty()
    {
        return Ok(formula_error(ScalarError::Value));
    }

    if function.is_sequence() {
        apply_sequence(evaluator, function, arguments)
    } else {
        apply_matrix(evaluator, function, arguments)
    }
}

fn apply_sequence<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    // SUMSQ and MULTINOMIAL have NumberSequence rather than
    // NumberSequenceList. A 3-D reference remains one Reference and is
    // therefore legal even when its physical representation contains
    // multiple sheet planes; only an explicit union/reference-list is
    // refused.
    if matches!(function, Function::SumSquares | Function::Multinomial)
        && arguments
            .iter()
            .any(|value| matches!(value, RuntimeValue::Areas(areas) if areas.is_list))
    {
        // A rejected ReferenceList remains unread, but a direct formula
        // error in a later argument still has source precedence.  Inspect
        // only materialized scalar/array values before publishing the shape
        // error; `direct_formula_error` deliberately never traverses Areas.
        if let Some(error) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    let mut fold = SequenceFold::new(function)?;
    for argument in arguments {
        evaluator.scalar.charge_work(1)?;
        match argument {
            RuntimeValue::Areas(areas) => {
                fold_reference_areas(evaluator, &mut fold, areas)?;
            },
            RuntimeValue::Array(array) => {
                for (index, element) in array.cells.into_iter().enumerate() {
                    evaluator.charge_cell_work(index)?;
                    fold_array_element(evaluator, &mut fold, element)?;
                }
            },
            RuntimeValue::ScalarCell(area) => {
                let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
                fold_scalar_value(evaluator, &mut fold, projected)?;
            },
            RuntimeValue::Empty => fold_scalar_number(evaluator, &mut fold, 0.0)?,
            RuntimeValue::Missing => fold.push_formula_error(ScalarError::Value),
            RuntimeValue::Scalar(value) => {
                fold_scalar_working(evaluator, &mut fold, value)?;
            },
            RuntimeValue::SourceReference => {
                return Err(EvaluationFailure::Unsupported(
                    super::super::UnsupportedKind::Reference,
                ));
            },
        }
    }
    fold.finish()
}

fn fold_reference_areas<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    areas: RuntimeAreaSet<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for area in areas.areas {
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                let index = fold.next_work_index()?;
                evaluator.charge_cell_work(index)?;
                let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                let element = evaluator.read_to_element(read)?;
                fold_reference_element(evaluator, fold, element)?;
            }
        }
    }
    Ok(())
}

fn fold_reference_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    element: RuntimeElement<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    // NumberSequence conversion includes only Number and Error members from a
    // reference. Empty, Text, and distinguished Logical cells are omitted.
    match element {
        RuntimeElement::Empty => {},
        RuntimeElement::Missing => fold.push_formula_error(ScalarError::Value),
        RuntimeElement::Present(WorkingValue::Number(value)) => {
            fold.push_number(evaluator, value)?;
        },
        RuntimeElement::Present(WorkingValue::Error(error)) => {
            fold.push_formula_error(error);
        },
        RuntimeElement::Present(WorkingValue::Logical(_))
        | RuntimeElement::Present(WorkingValue::Text(_))
        | RuntimeElement::Present(WorkingValue::Complex(_)) => {},
    }
    // Keep the mutable evaluator in this helper's signature so future
    // sequence conversion policies can charge text/provider work at the same
    // boundary without changing the streaming call sites.
    let _ = evaluator;
    Ok(())
}

fn fold_array_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    element: RuntimeElement<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty => fold_scalar_number(evaluator, fold, 0.0),
        RuntimeElement::Missing => {
            fold.push_formula_error(ScalarError::Value);
            Ok(())
        },
        RuntimeElement::Present(value) => fold_scalar_working(evaluator, fold, value),
    }
}

fn fold_scalar_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Empty => fold_scalar_number(evaluator, fold, 0.0),
        RuntimeValue::Missing => {
            fold.push_formula_error(ScalarError::Value);
            Ok(())
        },
        RuntimeValue::Scalar(value) => fold_scalar_working(evaluator, fold, value),
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            fold_scalar_value(evaluator, fold, projected)
        },
        RuntimeValue::Array(_) | RuntimeValue::Areas(_) => Err(
            EvaluationFailure::InvalidExpression("sequence scalar projection remained array-like"),
        ),
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn fold_scalar_working<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    value: WorkingValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        WorkingValue::Number(value) => fold.push_number(evaluator, value),
        WorkingValue::Logical(value) => fold_scalar_number(evaluator, fold, f64::from(value)),
        WorkingValue::Text(value) => {
            match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)? {
                Ok(value) => fold.push_number(evaluator, value),
                Err(error) => {
                    fold.push_generated_error(error);
                    Ok(())
                },
            }
        },
        WorkingValue::Error(error) => {
            fold.push_formula_error(error);
            Ok(())
        },
        WorkingValue::Complex(_) => {
            fold.push_generated_error(ScalarError::Value);
            Ok(())
        },
    }
}

fn fold_scalar_number<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut SequenceFold,
    value: f64,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    fold.push_number(evaluator, value)
}

struct SequenceFold {
    function: Function,
    accumulator: SequenceAccumulator,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
    work_index: usize,
}

enum SequenceAccumulator {
    Numeric(NumericAggregate),
    Squares(WideSum),
    Discrete(DiscreteFold),
}

impl SequenceFold {
    fn new(function: Function) -> EvaluationResult<Self> {
        let accumulator = match function {
            Function::Sum => {
                SequenceAccumulator::Numeric(NumericAggregate::new(NumericOperation::Sum))
            },
            Function::Product => {
                SequenceAccumulator::Numeric(NumericAggregate::new(NumericOperation::Product))
            },
            Function::SumSquares => SequenceAccumulator::Squares(WideSum::default()),
            Function::Gcd => {
                SequenceAccumulator::Discrete(DiscreteFold::new(DiscreteFunction::Gcd))
            },
            Function::Lcm => {
                SequenceAccumulator::Discrete(DiscreteFold::new(DiscreteFunction::Lcm))
            },
            Function::Multinomial => {
                SequenceAccumulator::Discrete(DiscreteFold::new(DiscreteFunction::Multinomial))
            },
            _ => {
                return Err(EvaluationFailure::InvalidExpression(
                    "matrix aggregate reached sequence reducer",
                ));
            },
        };
        Ok(Self {
            function,
            accumulator,
            formula_error: None,
            generated_error: None,
            work_index: 0,
        })
    }

    fn next_work_index(&mut self) -> EvaluationResult<usize> {
        let index = self.work_index;
        self.work_index =
            self.work_index
                .checked_add(1)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "aggregate sequence work index overflow",
                ))?;
        Ok(index)
    }

    fn push_number<'expr, 'scalar, 'exec, 'position, R>(
        &mut self,
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
        value: f64,
    ) -> EvaluationResult<()>
    where
        R: Resolver + ?Sized,
    {
        if self.formula_error.is_some() {
            return Ok(());
        }
        match &mut self.accumulator {
            SequenceAccumulator::Numeric(numeric) => {
                if let Err(error) = numeric.push_number(value) {
                    self.push_generated_error(error);
                }
            },
            SequenceAccumulator::Squares(squares) => {
                if let Err(error) = squares.push_square(value) {
                    self.push_generated_error(error);
                }
            },
            SequenceAccumulator::Discrete(discrete) => {
                // Admit bounded reducer work before executing its
                // fixed-width arithmetic, then mirror the first generated
                // reducer error into the outer sequence state.
                let work = discrete.work_for(value);
                evaluator
                    .scalar
                    .charge_work(u64::try_from(work).unwrap_or(u64::MAX))?;
                if let Err(error) = discrete.push_number(value) {
                    self.push_generated_error(error);
                }
            },
        }
        Ok(())
    }

    fn push_formula_error(&mut self, error: ScalarError) {
        if self.formula_error.is_none() {
            self.formula_error = Some(error);
        }
    }

    fn push_generated_error(&mut self, error: ScalarError) {
        if self.generated_error.is_none() {
            self.generated_error = Some(error);
        }
    }

    fn finish<'a>(self) -> EvaluationResult<RuntimeValue<'a>> {
        if let Some(error) = self.formula_error.or(self.generated_error) {
            return Ok(formula_error(error));
        }
        let value = match self.accumulator {
            SequenceAccumulator::Numeric(numeric) => match numeric.result() {
                Ok(Some(value)) => value,
                Ok(None) if self.function == Function::Product => 1.0,
                Ok(None) => 0.0,
                Err(error) => return Ok(formula_error(error)),
            },
            SequenceAccumulator::Squares(squares) => match squares.result() {
                Ok(Some(value)) => value,
                Ok(None) => 0.0,
                Err(error) => return Ok(formula_error(error)),
            },
            SequenceAccumulator::Discrete(discrete) => match discrete.finish() {
                Ok(value) => value,
                // A NumberSequenceList reference can contain no admitted
                // Number/Error cells after Empty/Text/Logical filtering. The
                // value profile publishes the constrained empty GCD/LCM as
                // #NUM!, while MULTINOMIAL keeps its empty product identity.
                Err(ScalarError::Value)
                    if matches!(self.function, Function::Gcd | Function::Lcm) =>
                {
                    return Ok(formula_error(ScalarError::Number));
                },
                Err(ScalarError::Value) if self.function == Function::Multinomial => 1.0,
                Err(error) => return Ok(formula_error(error)),
            },
        };
        Ok(number(value))
    }
}

fn apply_matrix<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(first) = arguments.first() else {
        return Ok(formula_error(ScalarError::Value));
    };
    let mut shape = matrix_shape(first).ok();
    let mut shape_error = matrix_shape(first).err();
    for argument in arguments.iter().skip(1) {
        match matrix_shape(argument) {
            Ok(candidate) if shape.is_some_and(|shape| candidate == shape) => {},
            Ok(_) | Err(_) => {
                shape = None;
                if shape_error.is_none() {
                    shape_error = Some(ScalarError::Value);
                }
            },
        }
    }
    let Some(shape) = shape else {
        // A direct array Error has source precedence over a shape failure;
        // inspect only the already-materialized argument values. References
        // are intentionally not read on this rejected shape path.
        if let Some(error) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(shape_error.unwrap_or(ScalarError::Value)));
    };
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "aggregate matrix cell count overflow",
        ))?;

    let mut fold = MatrixFold::new(function, arguments.len());
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        match function {
            Function::SumProduct => {
                if arguments.len() == 1 {
                    match matrix_number(evaluator, &arguments[0], shape, index)? {
                        Ok(value) if !fold.has_formula_error() => {
                            fold.push_single(index, value)?;
                        },
                        Ok(_) => {},
                        Err(error) => fold.push_input_error(0, index, error),
                    }
                } else if arguments.len() == 2 {
                    let left = matrix_number(evaluator, &arguments[0], shape, index)?;
                    let right = matrix_number(evaluator, &arguments[1], shape, index)?;
                    match (left, right) {
                        (Ok(left), Ok(right)) if !fold.has_formula_error() => {
                            fold.push_product_pair(index, left, right)?;
                        },
                        (left, right) => {
                            if let Err(error) = left {
                                fold.push_input_error(0, index, error);
                            }
                            if let Err(error) = right {
                                fold.push_input_error(1, index, error);
                            }
                        },
                    }
                } else {
                    let mut term = ProductTerm::default();
                    let mut valid = true;
                    for (argument_index, argument) in arguments.iter().enumerate() {
                        match matrix_number(evaluator, argument, shape, index)? {
                            Ok(value) => {
                                if term.push_factor(value).is_err() {
                                    valid = false;
                                    fold.push_generated_error(
                                        argument_index,
                                        index,
                                        ScalarError::Number,
                                    );
                                }
                            },
                            Err(error) => {
                                fold.push_input_error(argument_index, index, error);
                                valid = false;
                            },
                        }
                    }
                    if valid && !fold.has_formula_error() {
                        match fold.push_product_term(term)? {
                            Ok(()) => {},
                            Err(ProductSumError::Number) => fold.push_generated_error(
                                arguments.len(),
                                index,
                                ScalarError::Number,
                            ),
                            Err(ProductSumError::ExponentSpan { observed, limit }) => {
                                return Err(product_span_failure(evaluator, observed, limit));
                            },
                        }
                    }
                }
            },
            Function::SumX2My2 | Function::SumX2Py2 | Function::SumXMy2 => {
                let left = matrix_number(evaluator, &arguments[0], shape, index)?;
                let right = matrix_number(evaluator, &arguments[1], shape, index)?;
                match (left, right) {
                    (Ok(left), Ok(right)) if !fold.has_formula_error() => {
                        fold.push_pair(arguments.len(), index, function, left, right)?;
                    },
                    (left, right) => {
                        if let Err(error) = left {
                            fold.push_input_error(0, index, error);
                        }
                        if let Err(error) = right {
                            fold.push_input_error(1, index, error);
                        }
                    },
                }
            },
            Function::Sum
            | Function::Product
            | Function::SumSquares
            | Function::Gcd
            | Function::Lcm
            | Function::Multinomial => {
                return Err(EvaluationFailure::InvalidExpression(
                    "sequence aggregate reached matrix reducer",
                ));
            },
        }
    }
    fold.finish(evaluator)
}

fn matrix_shape(value: &RuntimeValue<'_>) -> Result<Shape, ScalarError> {
    match value {
        RuntimeValue::Array(array) => Ok(array.shape),
        RuntimeValue::Areas(areas) if areas.is_list || areas.areas.len() != 1 => {
            Err(ScalarError::Value)
        },
        RuntimeValue::Areas(areas) => {
            let area = areas.areas.first().ok_or(ScalarError::Value)?;
            Shape::new(area.rect.rows(), area.rect.columns()).map_err(|_| ScalarError::Value)
        },
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::ScalarCell(_)
        | RuntimeValue::SourceReference => Shape::new(1, 1).map_err(|_| ScalarError::Value),
    }
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
            | RuntimeValue::Areas(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::SourceReference => {},
        }
    }
    Ok(None)
}

fn product_span_failure<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    observed: u64,
    limit: u64,
) -> EvaluationFailure
where
    R: Resolver + ?Sized,
{
    // `ProductSumError::ExponentSpan` already reports the two-magnitude
    // fixed-profile admission in bytes. Keep those units intact when the
    // resolver-backed bridge turns it into a typed Memory refusal; converting
    // them a second time would under-report the required window.
    EvaluationFailure::ResourceLimit(evaluator.local_limit(
        Resource::Memory,
        observed,
        usize::try_from(limit).unwrap_or(usize::MAX),
    ))
}

fn matrix_number<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
    shape: Shape,
    index: usize,
) -> EvaluationResult<Result<f64, InputError>>
where
    R: Resolver + ?Sized,
{
    // Charge each argument-cell conversion separately. The enclosing matrix
    // loop accounts for the output position and cancellation cadence; this
    // charge covers the K physical inputs inspected at that position.
    evaluator.scalar.charge_work(1)?;
    let element = matrix_element(evaluator, value, shape, index)?;
    match element {
        RuntimeElement::Empty => Ok(Ok(0.0)),
        RuntimeElement::Missing => Ok(Err(InputError::Formula(ScalarError::Value))),
        RuntimeElement::Present(WorkingValue::Number(value)) if value.is_finite() => Ok(Ok(value)),
        RuntimeElement::Present(WorkingValue::Number(_)) => {
            Ok(Err(InputError::Generated(ScalarError::Number)))
        },
        RuntimeElement::Present(WorkingValue::Logical(value)) => Ok(Ok(f64::from(value))),
        RuntimeElement::Present(WorkingValue::Error(error)) => Ok(Err(InputError::Formula(error))),
        RuntimeElement::Present(WorkingValue::Text(value)) => {
            match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)? {
                Ok(value) => Ok(Ok(value)),
                Err(error) => Ok(Err(InputError::Generated(error))),
            }
        },
        RuntimeElement::Present(WorkingValue::Complex(_)) => {
            Ok(Err(InputError::Generated(ScalarError::Value)))
        },
    }
}

fn matrix_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
    shape: Shape,
    index: usize,
) -> EvaluationResult<RuntimeElement<'expr>>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Empty => Ok(RuntimeElement::Empty),
        RuntimeValue::Missing => Ok(RuntimeElement::Missing),
        RuntimeValue::Scalar(value) => Ok(RuntimeElement::Present(evaluator.clone_working(value)?)),
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(*area))?;
            matrix_element(evaluator, &projected, shape, index)
        },
        RuntimeValue::Array(array) => super::array_element_for(array, shape, index)
            .map(|element| evaluator.clone_element(element))
            .transpose()?
            .ok_or(EvaluationFailure::InvalidExpression(
                "aggregate array index is out of bounds",
            )),
        RuntimeValue::Areas(areas) => {
            let area = areas
                .areas
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "aggregate reference area is empty",
                ))?;
            let columns = area.rect.columns();
            let row = index / columns;
            let column = index % columns;
            if row >= area.rect.rows() || column >= columns {
                return Err(EvaluationFailure::InvalidExpression(
                    "aggregate reference index is out of bounds",
                ));
            }
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start + row,
                area.rect.column_start + column,
            )?;
            evaluator.read_to_element(read)
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

struct MatrixFold {
    accumulator: MatrixAccumulator,
    formula_error: Option<MatrixError>,
    generated_error: Option<MatrixError>,
}

enum MatrixAccumulator {
    Single(NumericAggregate),
    ProductPair(WideSum),
    Product(ScaledProductSum),
    Pair(WideSum),
}

#[derive(Clone, Copy)]
struct MatrixError {
    argument_index: usize,
    cell_index: usize,
    error: ScalarError,
}

impl MatrixError {
    fn precedes(self, other: Self) -> bool {
        (self.argument_index, self.cell_index) < (other.argument_index, other.cell_index)
    }
}

impl MatrixFold {
    fn new(function: Function, argument_count: usize) -> Self {
        let accumulator = if function == Function::SumProduct && argument_count == 1 {
            MatrixAccumulator::Single(NumericAggregate::new(NumericOperation::Sum))
        } else if function == Function::SumProduct && argument_count == 2 {
            MatrixAccumulator::ProductPair(WideSum::default())
        } else if function == Function::SumProduct {
            MatrixAccumulator::Product(ScaledProductSum::default())
        } else {
            MatrixAccumulator::Pair(WideSum::default())
        };
        Self {
            accumulator,
            formula_error: None,
            generated_error: None,
        }
    }

    fn push_product_term(
        &mut self,
        term: ProductTerm,
    ) -> EvaluationResult<Result<(), ProductSumError>> {
        let MatrixAccumulator::Product(sum) = &mut self.accumulator else {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix product accumulator is missing",
            ));
        };
        Ok(sum.push_term(term))
    }

    fn push_single(&mut self, cell_index: usize, value: f64) -> EvaluationResult<()> {
        let MatrixAccumulator::Single(sum) = &mut self.accumulator else {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix single-product accumulator is missing",
            ));
        };
        if let Err(error) = sum.push_number(value) {
            self.push_generated_error(0, cell_index, error);
        }
        Ok(())
    }

    fn push_product_pair(
        &mut self,
        cell_index: usize,
        left: f64,
        right: f64,
    ) -> EvaluationResult<()> {
        let MatrixAccumulator::ProductPair(sum) = &mut self.accumulator else {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix pair-product accumulator is missing",
            ));
        };
        if let Err(error) = sum.push_product(left, right) {
            self.push_generated_error(2, cell_index, error);
        }
        Ok(())
    }

    fn push_pair(
        &mut self,
        argument_index: usize,
        cell_index: usize,
        function: Function,
        left: f64,
        right: f64,
    ) -> EvaluationResult<()> {
        let MatrixAccumulator::Pair(sum) = &mut self.accumulator else {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix pair accumulator is missing",
            ));
        };
        let result = match function {
            Function::SumX2My2 => sum.push_difference_of_squares(left, right),
            Function::SumX2Py2 => sum.push_sum_of_squares(left, right),
            Function::SumXMy2 => sum.push_squared_difference(left, right),
            _ => {
                return Err(EvaluationFailure::InvalidExpression(
                    "non-pair function reached pair accumulator",
                ));
            },
        };
        if let Err(error) = result {
            self.push_generated_error(argument_index, cell_index, error);
        }
        Ok(())
    }

    fn push_input_error(&mut self, argument_index: usize, cell_index: usize, error: InputError) {
        match error {
            InputError::Formula(error) => {
                self.push_formula_error(argument_index, cell_index, error)
            },
            InputError::Generated(error) => {
                self.push_generated_error(argument_index, cell_index, error)
            },
        }
    }

    fn push_formula_error(&mut self, argument_index: usize, cell_index: usize, error: ScalarError) {
        let candidate = MatrixError {
            argument_index,
            cell_index,
            error,
        };
        if self
            .formula_error
            .is_none_or(|current| candidate.precedes(current))
        {
            self.formula_error = Some(candidate);
        }
    }

    fn push_generated_error(
        &mut self,
        argument_index: usize,
        cell_index: usize,
        error: ScalarError,
    ) {
        let candidate = MatrixError {
            argument_index,
            cell_index,
            error,
        };
        if self
            .generated_error
            .is_none_or(|current| candidate.precedes(current))
        {
            self.generated_error = Some(candidate);
        }
    }

    fn has_formula_error(&self) -> bool {
        self.formula_error.is_some()
    }

    fn finish<'expr, 'scalar, 'exec, 'position, R>(
        self,
        evaluator: &ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    ) -> EvaluationResult<RuntimeValue<'expr>>
    where
        R: Resolver + ?Sized,
    {
        if let Some(error) = self.formula_error.or(self.generated_error) {
            return Ok(formula_error(error.error));
        }
        let value = match self.accumulator {
            MatrixAccumulator::Single(sum) => match sum.result() {
                Ok(value) => value.unwrap_or(0.0),
                Err(error) => return Ok(formula_error(error)),
            },
            MatrixAccumulator::ProductPair(sum) => match sum.result() {
                Ok(value) => value.unwrap_or(0.0),
                Err(error) => return Ok(formula_error(error)),
            },
            MatrixAccumulator::Product(sum) => match sum.result() {
                Ok(value) => value.unwrap_or(0.0),
                Err(ProductSumError::Number) => return Ok(formula_error(ScalarError::Number)),
                Err(ProductSumError::ExponentSpan { observed, limit }) => {
                    return Err(product_span_failure(evaluator, observed, limit));
                },
            },
            MatrixAccumulator::Pair(sum) => match sum.result() {
                Ok(value) => value.unwrap_or(0.0),
                Err(error) => return Ok(formula_error(error)),
            },
        };
        Ok(number(value))
    }
}

#[derive(Clone, Copy)]
enum InputError {
    Formula(ScalarError),
    Generated(ScalarError),
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
