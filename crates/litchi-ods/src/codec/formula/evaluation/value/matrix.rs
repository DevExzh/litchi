//! Bounded numerical and shape-preserving matrix functions.
//!
//! The parent value evaluator owns dispatch and the value representation.  This
//! child keeps the matrix kernels together so that every temporary and result
//! buffer goes through the evaluator's shared storage admission path.

use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeArrayValue, RuntimeElement, RuntimeValue,
    ScalarError, Shape, ValueEvaluator, WorkingValue, ensure_capacity,
};
use litchi_core::{Reservation, Resource};

const MATRIX_SCRATCH_SCOPE: &str = "formula matrix arithmetic scratch";

/// Keep a retained vector before its accounting token so every early return
/// drops elements before releasing the corresponding reservation.
struct ReservedVec<T> {
    values: Vec<T>,
    reservation: Option<Reservation>,
}

/// Return whether `name` is one of the matrix functions implemented here.
pub(super) fn is_matrix_function(name: &str) -> bool {
    ["MDETERM", "MINVERSE", "MMULT", "MUNIT", "TRANSPOSE"]
        .iter()
        .any(|candidate| name.eq_ignore_ascii_case(candidate))
}

#[derive(Clone, Copy)]
struct NumericScan {
    formula_error: Option<ScalarError>,
    invalid_number: bool,
    invalid_value: bool,
}

impl NumericScan {
    const fn new() -> Self {
        Self {
            formula_error: None,
            invalid_number: false,
            invalid_value: false,
        }
    }

    fn merge(&mut self, other: Self) {
        if self.formula_error.is_none() {
            self.formula_error = other.formula_error;
        }
        self.invalid_number |= other.invalid_number;
        self.invalid_value |= other.invalid_value;
    }

    fn error(self) -> Option<ScalarError> {
        self.formula_error.or(if self.invalid_number {
            Some(ScalarError::Number)
        } else if self.invalid_value {
            Some(ScalarError::Value)
        } else {
            None
        })
    }
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

fn matrix_index(row: usize, column: usize, columns: usize) -> EvaluationResult<usize> {
    row.checked_mul(columns)
        .and_then(|index| index.checked_add(column))
        .ok_or(EvaluationFailure::InvalidExpression(
            "matrix index overflow",
        ))
}

fn checked_cell_count(array: &RuntimeArrayValue<'_>) -> EvaluationResult<usize> {
    let expected = array
        .shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "matrix cell count overflow",
        ))?;
    if array.cells.len() != expected {
        return Err(EvaluationFailure::InvalidExpression(
            "matrix storage length disagrees with shape",
        ));
    }
    Ok(expected)
}

fn is_reference_list(value: &RuntimeValue<'_>) -> bool {
    matches!(value, RuntimeValue::Areas(areas) if areas.is_list)
}

fn demote_reference_list_error<'a>(
    generated_reference_list_error: bool,
    value: RuntimeValue<'a>,
) -> RuntimeValue<'a> {
    if generated_reference_list_error
        && matches!(
            &value,
            RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Value))
        )
    {
        // `materialize_for_array` represents a reference list as a formula
        // Value error.  Keep that error catchable as Value, but mark it as a
        // generated invalid element while collecting multiple operands so an
        // existing Error in a later argument still takes precedence.
        RuntimeValue::Missing
    } else {
        value
    }
}

impl<'expr, 'scalar, 'exec, 'position, R> ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>
where
    R: Resolver + ?Sized,
{
    /// Apply a matrix function after the parent VM has evaluated its
    /// arguments in source order.  Formula errors are returned as values;
    /// cancellation, provider failures, and resource failures remain
    /// evaluator failures.
    pub(super) fn apply_matrix_values(
        &mut self,
        name: &str,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if !is_matrix_function(name) {
            return Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Function,
            ));
        }
        if name.eq_ignore_ascii_case("MDETERM") {
            return self.apply_mdeterminant(arguments);
        }
        if name.eq_ignore_ascii_case("MINVERSE") {
            return self.apply_minverse(arguments);
        }
        if name.eq_ignore_ascii_case("MMULT") {
            return self.apply_mmult(arguments);
        }
        if name.eq_ignore_ascii_case("MUNIT") {
            return self.apply_munit(arguments);
        }
        if name.eq_ignore_ascii_case("TRANSPOSE") {
            return self.apply_transpose(arguments);
        }
        Err(EvaluationFailure::InvalidExpression(
            "matrix function name dispatch is incomplete",
        ))
    }

    fn apply_mdeterminant(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if arguments.len() != 1 {
            return Ok(formula_error(ScalarError::Value));
        }
        let mut arguments = arguments.into_iter();
        let raw = arguments
            .next()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix argument missing",
            ))?;
        let generated_reference_list_error = is_reference_list(&raw);
        let value = demote_reference_list_error(
            generated_reference_list_error,
            self.materialize_for_array(raw)?,
        );
        if !generated_reference_list_error {
            if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &value {
                return Ok(formula_error(*error));
            }
        }
        let array = self.force_array(value)?;
        let scan = self.scan_numeric_array(&array)?;
        let shape = array.shape;
        if let Some(error) = scan.error() {
            return Ok(formula_error(error));
        }
        if shape.rows() != shape.columns() {
            return Ok(formula_error(ScalarError::Value));
        }
        let cells = checked_cell_count(&array)?;
        let mut values = self.new_scratch(cells, self.limits.max_array_cells)?;
        self.copy_numbers(&array, &mut values.values)?;
        let result = self.determinant(&mut values.values, shape.rows());
        drop(values);
        result
    }

    fn apply_minverse(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if arguments.len() != 1 {
            return Ok(formula_error(ScalarError::Value));
        }
        let mut arguments = arguments.into_iter();
        let raw = arguments
            .next()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix argument missing",
            ))?;
        let generated_reference_list_error = is_reference_list(&raw);
        let value = demote_reference_list_error(
            generated_reference_list_error,
            self.materialize_for_array(raw)?,
        );
        if !generated_reference_list_error {
            if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &value {
                return Ok(formula_error(*error));
            }
        }
        let array = self.force_array(value)?;
        let scan = self.scan_numeric_array(&array)?;
        let shape = array.shape;
        if let Some(error) = scan.error() {
            return Ok(formula_error(error));
        }
        if shape.rows() != shape.columns() {
            return Ok(formula_error(ScalarError::Value));
        }
        let cells = checked_cell_count(&array)?;
        let width = shape
            .columns()
            .checked_mul(2)
            .ok_or(EvaluationFailure::InvalidExpression(
                "inverse matrix width overflow",
            ))?;
        let augmented_cells =
            shape
                .rows()
                .checked_mul(width)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "inverse matrix storage overflow",
                ))?;
        let maximum = self
            .limits
            .max_array_cells
            .checked_mul(2)
            .unwrap_or(usize::MAX);
        let mut values = self.new_scratch(augmented_cells, maximum)?;
        self.copy_augmented(&array, &mut values.values, width)?;
        let inverse = self.inverse(&mut values.values, shape.rows(), width)?;
        match inverse {
            Ok(()) => {},
            Err(error) => {
                drop(values);
                return Ok(formula_error(error));
            },
        }
        let mut output = self.new_element_buffer(cells)?;
        for row in 0..shape.rows() {
            for column in 0..shape.columns() {
                let source = matrix_index(row, shape.columns() + column, width)?;
                let value = values.values.get(source).copied().ok_or(
                    EvaluationFailure::InvalidExpression("inverse matrix index outside storage"),
                )?;
                if !value.is_finite() {
                    drop(values);
                    return Ok(formula_error(ScalarError::Number));
                }
                output
                    .values
                    .push(RuntimeElement::Present(WorkingValue::Number(value)));
            }
        }
        drop(values);
        self.make_array(shape, output.values, output.reservation, None)
    }

    fn apply_mmult(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if arguments.len() != 2 {
            return Ok(formula_error(ScalarError::Value));
        }
        let mut arguments = arguments.into_iter();
        let left_raw = arguments
            .next()
            .ok_or(EvaluationFailure::InvalidExpression("left matrix missing"))?;
        let left_generated_reference_list_error = is_reference_list(&left_raw);
        let left_value = demote_reference_list_error(
            left_generated_reference_list_error,
            self.materialize_for_array(left_raw)?,
        );
        if !left_generated_reference_list_error {
            if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &left_value {
                return Ok(formula_error(*error));
            }
        }
        let left = self.force_array(left_value)?;
        let mut scan = self.scan_numeric_array(&left)?;
        let right_raw = arguments
            .next()
            .ok_or(EvaluationFailure::InvalidExpression("right matrix missing"))?;
        let right_generated_reference_list_error = is_reference_list(&right_raw);
        let right_value = demote_reference_list_error(
            right_generated_reference_list_error,
            self.materialize_for_array(right_raw)?,
        );
        if !right_generated_reference_list_error {
            if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &right_value {
                if let Some(left_error) = scan.formula_error {
                    return Ok(formula_error(left_error));
                }
                return Ok(formula_error(*error));
            }
        }
        let right = self.force_array(right_value)?;
        scan.merge(self.scan_numeric_array(&right)?);
        if let Some(error) = scan.error() {
            return Ok(formula_error(error));
        }
        if left.shape.columns() != right.shape.rows() {
            return Ok(formula_error(ScalarError::Value));
        }
        let output_shape = Shape::new(left.shape.rows(), right.shape.columns())?;
        let output_cells = output_shape.cell_count().ok_or_else(|| {
            EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::MAX,
                self.limits.max_array_cells,
            ))
        })?;
        let mut output = self.new_element_buffer(output_cells)?;
        for row in 0..output_shape.rows() {
            for column in 0..output_shape.columns() {
                let mut sum = 0.0;
                for index in 0..left.shape.columns() {
                    self.scalar.charge_work(1)?;
                    let left_index = matrix_index(row, index, left.shape.columns())?;
                    let right_index = matrix_index(index, column, right.shape.columns())?;
                    let left_value = numeric_element(left.cells.get(left_index).ok_or(
                        EvaluationFailure::InvalidExpression("left matrix index outside storage"),
                    )?)?;
                    let right_value = numeric_element(right.cells.get(right_index).ok_or(
                        EvaluationFailure::InvalidExpression("right matrix index outside storage"),
                    )?)?;
                    let product = left_value * right_value;
                    let next = sum + product;
                    if !product.is_finite() || !next.is_finite() {
                        return Ok(formula_error(ScalarError::Number));
                    }
                    sum = next;
                }
                output
                    .values
                    .push(RuntimeElement::Present(WorkingValue::Number(sum)));
            }
        }
        self.make_array(output_shape, output.values, output.reservation, None)
    }

    fn apply_munit(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if arguments.len() != 1 {
            return Ok(formula_error(ScalarError::Value));
        }
        let mut arguments = arguments.into_iter();
        let value = self.matrix_scalar_parameter(arguments.next().ok_or(
            EvaluationFailure::InvalidExpression("MUNIT argument missing"),
        )?)?;
        let number = match value {
            RuntimeValue::Scalar(WorkingValue::Number(number)) => number,
            RuntimeValue::Scalar(WorkingValue::Logical(value)) => {
                if value {
                    1.0
                } else {
                    0.0
                }
            },
            RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(formula_error(error)),
            RuntimeValue::Scalar(WorkingValue::Text(_))
            | RuntimeValue::Scalar(WorkingValue::Complex(_))
            | RuntimeValue::Missing => {
                return Ok(formula_error(ScalarError::Value));
            },
            RuntimeValue::Empty => 0.0,
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
                return Err(EvaluationFailure::InvalidExpression(
                    "MUNIT scalar projection remained array-like",
                ));
            },
        };
        if !number.is_finite() {
            return Ok(formula_error(ScalarError::Number));
        }
        let truncated = number.trunc();
        if truncated <= 0.0 {
            return Ok(formula_error(ScalarError::Value));
        }
        // Float-to-integer casts saturate at the machine bound.  Treat a
        // positive value outside `usize` as a host-size refusal instead of
        // silently turning it into `usize::MAX` and attempting a product.
        let usize_exclusive = 2.0_f64.powi(usize::BITS as i32);
        if truncated >= usize_exclusive {
            return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::MAX,
                self.limits.max_array_cells,
            )));
        }
        let size = truncated as usize;
        if size == 0 {
            return Ok(formula_error(ScalarError::Value));
        }
        let cells = match size.checked_mul(size) {
            Some(cells) => cells,
            None => {
                return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                    Resource::Objects,
                    u64::MAX,
                    self.limits.max_array_cells,
                )));
            },
        };
        let shape = Shape::new(size, size)?;
        let mut output = self.new_element_buffer(cells)?;
        for row in 0..size {
            for column in 0..size {
                self.charge_cell_work(matrix_index(row, column, size)?)?;
                let value = if row == column { 1.0 } else { 0.0 };
                output
                    .values
                    .push(RuntimeElement::Present(WorkingValue::Number(value)));
            }
        }
        self.make_array(shape, output.values, output.reservation, None)
    }

    fn apply_transpose(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if arguments.len() != 1 {
            return Ok(formula_error(ScalarError::Value));
        }
        let mut arguments = arguments.into_iter();
        let value = self.materialize_for_array(arguments.next().ok_or(
            EvaluationFailure::InvalidExpression("matrix argument missing"),
        )?)?;
        if let RuntimeValue::Scalar(WorkingValue::Error(error)) = &value {
            return Ok(formula_error(*error));
        }
        let array = self.force_array(value)?;
        let input_shape = array.shape;
        let input_cells = checked_cell_count(&array)?;
        let output_shape = Shape::new(input_shape.columns(), input_shape.rows())?;
        let mut output = self.new_element_buffer(input_cells)?;
        let RuntimeArrayValue {
            shape: _,
            cells,
            _reservation: input_reservation,
            origin: _,
            preserve_scalar_result: _,
        } = array;
        let mut input = ReservedVec {
            values: cells,
            reservation: input_reservation,
        };
        for row in 0..output_shape.rows() {
            for column in 0..output_shape.columns() {
                let source = matrix_index(column, row, input_shape.columns())?;
                let element = input
                    .values
                    .get_mut(source)
                    .map(|element| std::mem::replace(element, RuntimeElement::Empty))
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "transpose index outside storage",
                    ))?;
                self.charge_cell_work(matrix_index(row, column, output_shape.columns())?)?;
                output.values.push(element);
            }
        }
        drop(input);
        self.make_array(output_shape, output.values, output.reservation, None)
    }

    fn force_array(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeArrayValue<'expr>> {
        let value = self.materialize_for_array(value)?;
        match value {
            RuntimeValue::Array(array) => Ok(array),
            RuntimeValue::Empty => self.single_cell(RuntimeElement::Empty),
            RuntimeValue::Missing => self.single_cell(RuntimeElement::Missing),
            RuntimeValue::Scalar(value) => self.single_cell(RuntimeElement::Present(value)),
            RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => Err(
                EvaluationFailure::InvalidExpression("matrix argument remained a reference"),
            ),
        }
    }

    fn single_cell(
        &mut self,
        element: RuntimeElement<'expr>,
    ) -> EvaluationResult<RuntimeArrayValue<'expr>> {
        let shape = Shape::new(1, 1)?;
        let mut cells = self.new_element_buffer(1)?;
        cells.values.push(element);
        Ok(RuntimeArrayValue {
            shape,
            cells: cells.values,
            _reservation: cells.reservation,
            origin: None,
            preserve_scalar_result: false,
        })
    }

    fn scan_numeric_array(
        &mut self,
        array: &RuntimeArrayValue<'expr>,
    ) -> EvaluationResult<NumericScan> {
        let _ = checked_cell_count(array)?;
        let mut scan = NumericScan::new();
        for (index, element) in array.cells.iter().enumerate() {
            self.charge_cell_work(index)?;
            match element {
                RuntimeElement::Present(WorkingValue::Number(value)) => {
                    if !value.is_finite() {
                        scan.invalid_number = true;
                    }
                },
                RuntimeElement::Present(WorkingValue::Error(error)) => {
                    if scan.formula_error.is_none() {
                        scan.formula_error = Some(*error);
                    }
                },
                RuntimeElement::Empty
                | RuntimeElement::Missing
                | RuntimeElement::Present(WorkingValue::Logical(_))
                | RuntimeElement::Present(WorkingValue::Text(_))
                | RuntimeElement::Present(WorkingValue::Complex(_)) => {
                    scan.invalid_value = true;
                },
            }
        }
        Ok(scan)
    }

    fn copy_numbers(
        &mut self,
        array: &RuntimeArrayValue<'expr>,
        output: &mut Vec<f64>,
    ) -> EvaluationResult<()> {
        for (index, element) in array.cells.iter().enumerate() {
            self.charge_cell_work(index)?;
            let value = numeric_element(element)?;
            if !value.is_finite() {
                return Err(EvaluationFailure::InvalidExpression(
                    "non-finite numeric matrix escaped validation",
                ));
            }
            output.push(value);
        }
        Ok(())
    }

    fn copy_augmented(
        &mut self,
        array: &RuntimeArrayValue<'expr>,
        output: &mut Vec<f64>,
        width: usize,
    ) -> EvaluationResult<()> {
        let n = array.shape.rows();
        let columns = array.shape.columns();
        for row in 0..n {
            for column in 0..width {
                let target = matrix_index(row, column, width)?;
                let value = if column < columns {
                    let source = matrix_index(row, column, columns)?;
                    numeric_element(array.cells.get(source).ok_or(
                        EvaluationFailure::InvalidExpression(
                            "inverse source index outside storage",
                        ),
                    )?)?
                } else if column - columns == row {
                    1.0
                } else {
                    0.0
                };
                if !value.is_finite() {
                    return Err(EvaluationFailure::InvalidExpression(
                        "non-finite numeric matrix escaped validation",
                    ));
                }
                self.charge_cell_work(target)?;
                output.push(value);
            }
        }
        Ok(())
    }

    fn new_scratch(&mut self, cells: usize, maximum: usize) -> EvaluationResult<ReservedVec<f64>> {
        let mut values = Vec::new();
        let mut reservation = None;
        ensure_capacity(
            &mut values,
            &mut reservation,
            cells,
            maximum,
            self.execution,
            &self.storage_budget,
            MATRIX_SCRATCH_SCOPE,
        )?;
        Ok(ReservedVec {
            values,
            reservation,
        })
    }

    fn new_element_buffer(
        &mut self,
        cells: usize,
    ) -> EvaluationResult<ReservedVec<RuntimeElement<'expr>>> {
        let (values, reservation) = self.new_element_vec(cells)?;
        Ok(ReservedVec {
            values,
            reservation,
        })
    }

    fn determinant(
        &mut self,
        values: &mut [f64],
        size: usize,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let mut determinant = 1.0;
        let mut sign = 1.0;
        for column in 0..size {
            let pivot = match self.choose_pivot(values, size, size, column, column)? {
                Some(pivot) => pivot,
                None => return Ok(RuntimeValue::Scalar(WorkingValue::Number(0.0))),
            };
            if pivot != column {
                self.swap_rows(values, size, size, pivot, column)?;
                sign = -sign;
            }
            let pivot_index = matrix_index(column, column, size)?;
            let pivot_value =
                *values
                    .get(pivot_index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "determinant pivot outside storage",
                    ))?;
            if !pivot_value.is_finite() {
                return Ok(formula_error(ScalarError::Number));
            }
            determinant *= pivot_value;
            if !determinant.is_finite() {
                return Ok(formula_error(ScalarError::Number));
            }
            for row in column + 1..size {
                self.scalar.charge_work(1)?;
                let row_pivot = matrix_index(row, column, size)?;
                let factor = *values
                    .get(row_pivot)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "determinant row outside storage",
                    ))?
                    / pivot_value;
                if !factor.is_finite() {
                    return Ok(formula_error(ScalarError::Number));
                }
                if let Some(value) = values.get_mut(row_pivot) {
                    *value = 0.0;
                }
                for trailing in column + 1..size {
                    self.scalar.charge_work(1)?;
                    let target = matrix_index(row, trailing, size)?;
                    let source = matrix_index(column, trailing, size)?;
                    let pivot_entry =
                        *values
                            .get(source)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "determinant pivot row outside storage",
                            ))?;
                    let entry =
                        values
                            .get_mut(target)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "determinant target outside storage",
                            ))?;
                    *entry -= factor * pivot_entry;
                    if !entry.is_finite() {
                        return Ok(formula_error(ScalarError::Number));
                    }
                }
            }
        }
        determinant *= sign;
        if determinant.is_finite() {
            Ok(RuntimeValue::Scalar(WorkingValue::Number(determinant)))
        } else {
            Ok(formula_error(ScalarError::Number))
        }
    }

    fn inverse(
        &mut self,
        values: &mut [f64],
        size: usize,
        width: usize,
    ) -> EvaluationResult<Result<(), ScalarError>> {
        for column in 0..size {
            let pivot = match self.choose_pivot(values, size, width, column, column)? {
                Some(pivot) => pivot,
                None => return Ok(Err(ScalarError::Number)),
            };
            if pivot != column {
                self.swap_rows(values, width, width, pivot, column)?;
            }
            let pivot_index = matrix_index(column, column, width)?;
            let pivot_value =
                *values
                    .get(pivot_index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "inverse pivot outside storage",
                    ))?;
            if !pivot_value.is_finite() || pivot_value == 0.0 {
                return Ok(Err(ScalarError::Number));
            }
            for entry in 0..width {
                self.scalar.charge_work(1)?;
                let index = matrix_index(column, entry, width)?;
                let value = values
                    .get_mut(index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "inverse pivot row outside storage",
                    ))?;
                *value /= pivot_value;
                if !value.is_finite() {
                    return Ok(Err(ScalarError::Number));
                }
            }
            for row in 0..size {
                if row == column {
                    continue;
                }
                let factor_index = matrix_index(row, column, width)?;
                let factor =
                    *values
                        .get(factor_index)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "inverse row outside storage",
                        ))?;
                self.scalar.charge_work(1)?;
                if !factor.is_finite() {
                    return Ok(Err(ScalarError::Number));
                }
                for entry in 0..width {
                    self.scalar.charge_work(1)?;
                    let target = matrix_index(row, entry, width)?;
                    let source = matrix_index(column, entry, width)?;
                    let pivot_entry =
                        *values
                            .get(source)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "inverse pivot row outside storage",
                            ))?;
                    let value =
                        values
                            .get_mut(target)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "inverse target outside storage",
                            ))?;
                    *value -= factor * pivot_entry;
                    if !value.is_finite() {
                        return Ok(Err(ScalarError::Number));
                    }
                }
            }
        }
        Ok(Ok(()))
    }

    fn choose_pivot(
        &mut self,
        values: &[f64],
        rows: usize,
        width: usize,
        column: usize,
        first_row: usize,
    ) -> EvaluationResult<Option<usize>> {
        let mut best = None;
        let mut best_magnitude = 0.0;
        for row in first_row..rows {
            self.scalar.charge_work(1)?;
            let index = matrix_index(row, column, width)?;
            let value = *values
                .get(index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "matrix pivot outside storage",
                ))?;
            if !value.is_finite() {
                return Err(EvaluationFailure::InvalidExpression(
                    "non-finite matrix value escaped validation",
                ));
            }
            let magnitude = value.abs();
            if magnitude > best_magnitude {
                best_magnitude = magnitude;
                best = Some(row);
            }
        }
        Ok(best)
    }

    fn swap_rows(
        &mut self,
        values: &mut [f64],
        row_width: usize,
        storage_width: usize,
        first: usize,
        second: usize,
    ) -> EvaluationResult<()> {
        for column in 0..row_width {
            self.scalar.charge_work(1)?;
            let first_index = matrix_index(first, column, storage_width)?;
            let second_index = matrix_index(second, column, storage_width)?;
            let first_value =
                *values
                    .get(first_index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "matrix row outside storage",
                    ))?;
            let second_value =
                *values
                    .get(second_index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "matrix row outside storage",
                    ))?;
            *values
                .get_mut(first_index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "matrix row outside storage",
                ))? = second_value;
            *values
                .get_mut(second_index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "matrix row outside storage",
                ))? = first_value;
        }
        Ok(())
    }
}

fn numeric_element(element: &RuntimeElement<'_>) -> EvaluationResult<f64> {
    match element {
        RuntimeElement::Present(WorkingValue::Number(value)) if value.is_finite() => Ok(*value),
        _ => Err(EvaluationFailure::InvalidExpression(
            "non-numeric matrix element escaped validation",
        )),
    }
}
