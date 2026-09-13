//! Value-evaluator support for the complex-number sequence functions.
//!
//! Scalar complex kernels live in the parent evaluation module.  This child
//! only supplies the resolver-aware part of `IMSUM` and `IMPRODUCT`: it walks
//! arrays and resolved reference areas in source order and folds each element
//! as it is seen.  No temporary sequence is retained, so a large range is
//! bounded by the caller's reference-read, work, and storage limits.

use super::super::complex::Complex;
use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeElement, RuntimeValue, ScalarError,
    ValueEvaluator, WorkingValue,
};

/// The two section 6.8 functions whose parameters are `ComplexSequence`.
pub(super) fn is_complex_sequence_function(name: &str) -> bool {
    name.eq_ignore_ascii_case("IMSUM") || name.eq_ignore_ascii_case("IMPRODUCT")
}

struct Fold {
    product: bool,
    result: Complex,
    /// Formula errors already present in an operand. §4.6 gives these
    /// precedence over a conversion/domain failure discovered while the
    /// remaining sequence is scanned.
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
    work_index: usize,
}

impl Fold {
    fn new(product: bool) -> Result<Self, EvaluationFailure> {
        let result = Complex::new(if product { 1.0 } else { 0.0 }, 0.0, 'i').map_err(|_| {
            EvaluationFailure::InvalidExpression("complex sequence identity is not finite")
        })?;
        Ok(Self {
            product,
            result,
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
                    "complex sequence work index overflow",
                ))?;
        Ok(index)
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

    fn push_complex(&mut self, value: Complex) {
        if self.formula_error.is_some() || self.generated_error.is_some() {
            return;
        }
        let left = self.result;
        let result = if self.product {
            super::super::complex::complex_mul(left, value)
        } else {
            super::super::complex::complex_add(left, value)
        };
        match result {
            Ok(value) => self.result = value,
            Err(error) => self.push_generated_error(error),
        }
    }
}

/// Apply one of the two resolver-aware complex sequence functions.
pub(super) fn apply_sequence<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let product = name.eq_ignore_ascii_case("IMPRODUCT");
    if !is_complex_sequence_function(name) {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    }
    if arguments.is_empty() {
        return if product {
            Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Value,
            )))
        } else {
            Ok(RuntimeValue::Scalar(WorkingValue::Number(0.0)))
        };
    }

    let mut fold = Fold::new(product)?;
    for argument in arguments {
        evaluator.scalar.charge_work(1)?;
        match argument {
            RuntimeValue::Areas(areas) => {
                // A ReferenceList is processed record by record, and each
                // cuboid is processed sheet, row, then column.  The standard
                // permits row- or column-major order inside a sheet; this
                // profile chooses row-major deterministically.
                for area in areas.areas {
                    for row in area.rect.row_start..area.rect.row_end {
                        for column in area.rect.column_start..area.rect.column_end {
                            let index = fold.next_work_index()?;
                            evaluator.charge_cell_work(index)?;
                            let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                            let element = evaluator.read_to_element(read)?;
                            fold_reference_element(evaluator, &mut fold, element)?;
                        }
                    }
                }
            },
            RuntimeValue::Array(array) => {
                // Inline arrays are value sequences.  Their Empty and Logical
                // members use the scalar conversion profile, while a
                // reference area below follows §6.3.11's omission rules.
                for element in array.cells {
                    let index = fold.next_work_index()?;
                    evaluator.charge_cell_work(index)?;
                    fold_array_element(evaluator, &mut fold, element)?;
                }
            },
            RuntimeValue::Scalar(value) => {
                fold_scalar_value(evaluator, &mut fold, value)?;
            },
            RuntimeValue::ScalarCell(area) => {
                // `ScalarCell` is emitted only by scalar-demand arithmetic
                // frames.  Generic function arguments deliberately use
                // `VisitArgument`/`VisitMatrixArgument`, which retain a
                // reference as `Areas` and therefore reach the
                // §6.3.11 omission path above.  If a future scheduler sends
                // this token here, it has already selected the reference's
                // one-cell scalar profile, so Logical/Empty conversion is
                // intentional and cannot be confused with a reference
                // sequence element.
                let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
                match projected {
                    RuntimeValue::Scalar(value) => {
                        fold_scalar_value(evaluator, &mut fold, value)?;
                    },
                    RuntimeValue::Empty => fold_empty(&mut fold),
                    RuntimeValue::Missing => fold.push_formula_error(ScalarError::Value),
                    RuntimeValue::ScalarCell(_)
                    | RuntimeValue::Array(_)
                    | RuntimeValue::Areas(_) => {
                        return Err(EvaluationFailure::InvalidExpression(
                            "complex sequence scalar projection remained non-scalar",
                        ));
                    },
                }
            },
            // An Empty value outside a reference is converted through the
            // existing numeric Empty profile (zero).  Empty cells in a
            // resolved reference are handled by `fold_reference_element` and
            // omitted as required by ComplexSequence.
            RuntimeValue::Empty => fold_empty(&mut fold),
            RuntimeValue::Missing => fold.push_formula_error(ScalarError::Value),
        }
    }

    Ok(RuntimeValue::Scalar(
        match fold.formula_error.or(fold.generated_error) {
            Some(error) => WorkingValue::Error(error),
            None => WorkingValue::Complex(fold.result),
        },
    ))
}

fn fold_empty(fold: &mut Fold) {
    match Complex::new(0.0, 0.0, 'i') {
        Ok(value) => fold.push_complex(value),
        Err(_) => fold.push_generated_error(ScalarError::Number),
    }
}

fn fold_scalar_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut Fold,
    value: WorkingValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    let is_text = matches!(&value, WorkingValue::Text(_));
    let is_formula_error = matches!(&value, WorkingValue::Error(_));
    let converted = super::super::complex::to_complex(value, &mut evaluator.scalar)?;
    match converted {
        Ok(value) => fold.push_complex(value),
        Err(_error) if !fold.product && is_text => {
            // IMSUM explicitly ignores text that cannot be converted.  This
            // includes the Number error used for a non-finite parsed value.
        },
        Err(error) if is_formula_error => fold.push_formula_error(error),
        Err(error) => fold.push_generated_error(error),
    }
    Ok(())
}

fn fold_reference_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut Fold,
    element: RuntimeElement<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty => {},
        RuntimeElement::Missing => fold.push_formula_error(ScalarError::Value),
        RuntimeElement::Present(WorkingValue::Logical(_)) => {},
        RuntimeElement::Present(value) => fold_scalar_value(evaluator, fold, value)?,
    }
    Ok(())
}

fn fold_array_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    fold: &mut Fold,
    element: RuntimeElement<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty => fold_empty(fold),
        RuntimeElement::Missing => fold.push_formula_error(ScalarError::Value),
        RuntimeElement::Present(value) => fold_scalar_value(evaluator, fold, value)?,
    }
    Ok(())
}
