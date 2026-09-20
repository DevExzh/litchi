//! Resolver-backed HLOOKUP, VLOOKUP, LOOKUP, and MATCH search kernels.
//!
//! The lookup family keeps its data source as either a borrowed array view or
//! one resolved rectangular area.  Search cells are read on demand, so an
//! exact search reads in source order until its first match and an ordered
//! search uses a bounded binary search.  Only the selected result cell is read
//! after a match; no range-sized temporary vector is built here.

use super::super::super::lookup::{self as scalar_lookup, Function};
use super::super::{SheetExtent, SheetRef};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeArea, RuntimeArrayValue, RuntimeElement,
    RuntimeValue, ScalarError, Shape, ValueEvaluator, WorkingValue,
};
use crate::codec::formula::evaluation::value::CellRead;

fn is_search_function_kind(function: Function) -> bool {
    matches!(
        function,
        Function::HLookup | Function::Lookup | Function::Match | Function::VLookup
    )
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Axis {
    Row,
    Column,
}

impl Axis {
    fn length(self, shape: Shape) -> usize {
        match self {
            Self::Row => shape.columns(),
            Self::Column => shape.rows(),
        }
    }

    fn coordinate(self, index: usize) -> (usize, usize) {
        match self {
            Self::Row => (0, index),
            Self::Column => (index, 0),
        }
    }
}

#[derive(Clone, Copy)]
enum SearchSource<'value, 'expr> {
    Array(&'value RuntimeArrayValue<'expr>),
    Area(RuntimeArea<'expr>),
}

#[derive(Clone, Copy)]
struct SearchGrid<'value, 'expr> {
    source: SearchSource<'value, 'expr>,
    shape: Shape,
    /// A direct one-cell Reference result may be extended by LOOKUP. Arrays
    /// never receive this extension, even when they have one cell.
    reference: Option<RuntimeArea<'expr>>,
}

impl<'value, 'expr> SearchGrid<'value, 'expr> {
    fn vector_axis(self) -> Result<Axis, ScalarError> {
        match (self.shape.rows(), self.shape.columns()) {
            (1, _) => Ok(Axis::Row),
            (_, 1) => Ok(Axis::Column),
            _ => Err(ScalarError::Value),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum SearchCell<'a> {
    Empty,
    Missing,
    Number(f64),
    Logical(bool),
    Text(&'a str),
    Formula(ScalarError),
    Generated(ScalarError),
    Complex,
}

#[derive(Clone, Copy, Debug)]
enum SearchError {
    Formula(ScalarError),
    Generated(ScalarError),
}

impl SearchError {
    fn into_runtime<'a>(self) -> RuntimeValue<'a> {
        formula_error(match self {
            Self::Formula(error) | Self::Generated(error) => error,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MatchKind {
    Exact,
    Ascending,
    Descending,
}

#[derive(Clone, Copy, Debug)]
struct MatchResult {
    index: usize,
}

/// Apply one search function after the parent VM has visited all arguments.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = Function::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        ));
    };
    if !is_search_function_kind(function) {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        ));
    }
    if arguments
        .iter()
        .any(|value| matches!(value, RuntimeValue::SourceReference))
    {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        ));
    }
    if arguments
        .iter()
        .any(|value| matches!(value, RuntimeValue::Areas(areas) if areas.is_list))
    {
        if let Some(error) = first_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    match function {
        Function::HLookup => apply_table_lookup(evaluator, node, arguments, Axis::Row),
        Function::VLookup => apply_table_lookup(evaluator, node, arguments, Axis::Column),
        Function::Lookup => apply_lookup(evaluator, node, arguments),
        Function::Match => apply_match(evaluator, node, arguments),
        _ => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        )),
    }
}

/// Apply one search with borrowed argument slots.
///
/// Matrix lookup lifting evaluates the scalar key and selector arguments for
/// one output coordinate, while retaining the ForceArray data descriptors for
/// all coordinates.  Taking references here keeps a reference range or an
/// inline data array out of the per-output allocation path.  The ordinary
/// `apply` entry point above remains the ownership-taking scalar path.
pub(super) fn apply_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    name: &str,
    arguments: &[&RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = Function::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        ));
    };
    if !is_search_function_kind(function) {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        ));
    }
    if arguments
        .iter()
        .any(|value| matches!(*value, &RuntimeValue::SourceReference))
    {
        return Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        ));
    }
    if arguments
        .iter()
        .any(|value| matches!(*value, RuntimeValue::Areas(areas) if areas.is_list))
    {
        if let Some(error) = first_formula_error_borrowed(evaluator, arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    match function {
        Function::HLookup => apply_table_lookup_borrowed(evaluator, node, arguments, Axis::Row),
        Function::VLookup => apply_table_lookup_borrowed(evaluator, node, arguments, Axis::Column),
        Function::Lookup => apply_lookup_borrowed(evaluator, node, arguments),
        Function::Match => apply_match_borrowed(evaluator, node, arguments),
        _ => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Function,
        )),
    }
}

/// Validate ForceArray search inputs before a lifted key or selector is
/// projected.  Descriptor shape/list/type refusal is metadata-only and must
/// therefore happen before a scalar reference key can trigger a provider read.
pub(super) fn preflight<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<Result<(), ScalarError>>
where
    R: Resolver + ?Sized,
{
    let data = match function {
        Function::HLookup | Function::VLookup | Function::Match => arguments.get(1),
        Function::Lookup => arguments.get(1),
        _ => None,
    };
    let Some(data) = data else {
        return Ok(Err(ScalarError::Value));
    };
    let data_grid = match grid(evaluator, data)? {
        Ok(grid) => grid,
        Err(error) => return Ok(Err(error)),
    };
    if function == Function::Match && data_grid.vector_axis().is_err() {
        return Ok(Err(ScalarError::Value));
    }
    if function == Function::Lookup && arguments.len() == 3 {
        let Some(results) = arguments.get(2) else {
            return Ok(Err(ScalarError::Value));
        };
        if !matches!(results, &RuntimeValue::Missing) {
            let result = match grid(evaluator, results)? {
                Ok(grid) => grid,
                Err(error) => return Ok(Err(error)),
            };
            if result.vector_axis().is_err() {
                return Ok(Err(ScalarError::Value));
            }
        }
    }
    Ok(Ok(()))
}

fn apply_table_lookup_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: &[&RuntimeValue<'expr>],
    axis: Axis,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(3..=4).contains(&arguments.len()) {
        return invalid_arity_borrowed(evaluator, node, arguments);
    }
    // Validate scalar selectors before converting the lookup key. The key may
    // be a reference whose implicit intersection reads a provider cell.
    let lookup_error = formula_error_of(arguments[0]);
    let data = arguments[1];
    let grid = match grid(evaluator, data)? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let dimension = match axis {
        Axis::Row => grid.shape.rows(),
        Axis::Column => grid.shape.columns(),
    };
    let (index, range) = match table_controls(
        evaluator,
        arguments[2],
        arguments.get(3).copied(),
        dimension,
    )? {
        Ok(controls) => controls,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let search_len = axis.length(grid.shape);
    if search_len == 0 {
        return Ok(formula_error(
            lookup_error.unwrap_or(ScalarError::NotAvailable),
        ));
    }
    let lookup = scalar_argument_borrowed(evaluator, arguments[0])?;
    finish_table_lookup(evaluator, grid, axis, index, range, lookup)
}

fn finish_table_lookup<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'_, 'expr>,
    axis: Axis,
    index: usize,
    range: MatchKind,
    lookup: RuntimeValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let find = match key_from_runtime(evaluator, &lookup)? {
        Ok(key) => key,
        Err(error) => return Ok(error.into_runtime()),
    };
    let matched = match find_match(evaluator, grid, axis, find, range)? {
        Ok(result) => result,
        Err(error) => return Ok(error.into_runtime()),
    };
    let (row, column) = match axis {
        Axis::Row => (index - 1, matched.index),
        Axis::Column => (matched.index, index - 1),
    };
    selected_value(evaluator, grid, row, column)
}

fn apply_lookup_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: &[&RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(2..=3).contains(&arguments.len()) {
        return invalid_arity_borrowed(evaluator, node, arguments);
    }
    let lookup = scalar_argument_borrowed(evaluator, arguments[0])?;
    let lookup_error = formula_error_of(&lookup);
    let searched = match grid(evaluator, arguments[1])? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let search_axis = match searched.vector_axis() {
        Ok(axis) => axis,
        Err(_) => {
            if searched.shape.rows() >= searched.shape.columns() {
                Axis::Column
            } else {
                Axis::Row
            }
        },
    };
    let result_grid = if arguments.len() == 3 && !matches!(arguments[2], &RuntimeValue::Missing) {
        let grid = match grid(evaluator, arguments[2])? {
            Ok(grid) => grid,
            Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
        };
        if grid.vector_axis().is_err() {
            return Ok(formula_error(lookup_error.unwrap_or(ScalarError::Value)));
        }
        Some(grid)
    } else {
        None
    };

    let find = match key_from_runtime(evaluator, &lookup)? {
        Ok(key) => key,
        Err(error) => return Ok(error.into_runtime()),
    };
    let matched = match find_match(evaluator, searched, search_axis, find, MatchKind::Ascending)? {
        Ok(result) => result,
        Err(error) => return Ok(error.into_runtime()),
    };
    if let Some(result_grid) = result_grid {
        let Some(result_axis) = result_grid.vector_axis().ok() else {
            return Ok(formula_error(ScalarError::Value));
        };
        let result_length = result_axis.length(result_grid.shape);
        let result_grid = if matched.index >= result_length {
            match extend_result_reference(
                evaluator,
                result_grid,
                matched.index,
                search_axis.length(searched.shape),
            )? {
                Ok(grid) => grid,
                Err(error) => return Ok(formula_error(error)),
            }
        } else {
            result_grid
        };
        let (row, column) = match result_coordinate(result_grid, matched.index) {
            Ok(coordinate) => coordinate,
            Err(error) => return Ok(formula_error(error)),
        };
        return selected_value(evaluator, result_grid, row, column);
    }
    let (row, column) = match search_axis {
        Axis::Column => (matched.index, searched.shape.columns() - 1),
        Axis::Row => (searched.shape.rows() - 1, matched.index),
    };
    selected_value(evaluator, searched, row, column)
}

fn apply_match_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: &[&RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(2..=3).contains(&arguments.len()) {
        return invalid_arity_borrowed(evaluator, node, arguments);
    }
    // MATCH's mode is a scalar selector. Resolve/refuse it before the lookup
    // key's implicit intersection can read a reference cell.
    let lookup_error = formula_error_of(arguments[0]);
    let match_kind = if arguments.len() == 3 {
        match match_type_borrowed(evaluator, arguments[2])? {
            Ok(kind) => kind,
            Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
        }
    } else {
        MatchKind::Ascending
    };
    let grid = match grid(evaluator, arguments[1])? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let axis = match grid.vector_axis() {
        Ok(axis) => axis,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let lookup = scalar_argument_borrowed(evaluator, arguments[0])?;
    let find = match key_from_runtime(evaluator, &lookup)? {
        Ok(key) => key,
        Err(error) => return Ok(error.into_runtime()),
    };
    let matched = match find_match(evaluator, grid, axis, find, match_kind)? {
        Ok(result) => result,
        Err(error) => return Ok(error.into_runtime()),
    };
    Ok(RuntimeValue::Scalar(WorkingValue::Number(
        matched.index as f64 + 1.0,
    )))
}

fn invalid_arity_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: &[&RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if let Some(error) = first_formula_error_borrowed(evaluator, arguments)? {
        return Ok(formula_error(error));
    }
    if node.child_count() != 0 {
        evaluator.scalar.charge_work(node.child_count() as u64)?;
    }
    Ok(formula_error(ScalarError::Value))
}

fn first_formula_error_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    arguments: &[&RuntimeValue<'expr>],
) -> EvaluationResult<Option<ScalarError>>
where
    R: Resolver + ?Sized,
{
    for argument in arguments {
        match &**argument {
            RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(Some(*error)),
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
            | RuntimeValue::Areas(_)
            | RuntimeValue::SourceReference => {},
        }
    }
    Ok(None)
}

fn scalar_argument_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if evaluator.mode == super::super::Mode::Scalar && evaluator.projection.is_none() {
        return match value {
            RuntimeValue::Array(array) => {
                let element = evaluator.project_array_element(array)?;
                evaluator.element_to_runtime(element)
            },
            RuntimeValue::Areas(areas) => evaluator.project_area_value(areas),
            RuntimeValue::ScalarCell(area) => evaluator.project_scalar_cell(*area),
            RuntimeValue::Scalar(value) => evaluator.clone_working(value).map(RuntimeValue::Scalar),
            RuntimeValue::Empty => Ok(RuntimeValue::Empty),
            RuntimeValue::Missing => Ok(RuntimeValue::Missing),
            RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
                super::super::super::UnsupportedKind::Reference,
            )),
        };
    }
    match value {
        RuntimeValue::Array(array) => {
            let element = array
                .cells
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "empty matrix parameter",
                ))?;
            let element = evaluator.clone_element(element)?;
            evaluator.element_to_runtime(element)
        },
        RuntimeValue::Areas(areas) if areas.is_list => Ok(formula_error(ScalarError::Value)),
        RuntimeValue::Areas(areas) => {
            if areas.areas.len() != 1 {
                return Err(EvaluationFailure::Unsupported(
                    super::super::super::UnsupportedKind::Reference,
                ));
            }
            let area = &areas.areas[0];
            evaluator.scalar.charge_work(1)?;
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start,
                area.rect.column_start,
            )?;
            let element = evaluator.read_to_element(read)?;
            evaluator.element_to_runtime(element)
        },
        RuntimeValue::Scalar(value) => evaluator.clone_working(value).map(RuntimeValue::Scalar),
        RuntimeValue::Empty => Ok(RuntimeValue::Empty),
        RuntimeValue::Missing => Ok(RuntimeValue::Missing),
        RuntimeValue::ScalarCell(area) => {
            evaluator.charge_cell_work(0)?;
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start,
                area.rect.column_start,
            )?;
            let element = evaluator.read_to_element(read)?;
            evaluator.element_to_runtime(element)
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        )),
    }
}

/// Validate available controls before resolving reference-valued controls.
/// A known refusal must not fetch unrelated keys or selectors.
fn table_controls<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    index: &RuntimeValue<'expr>,
    range: Option<&RuntimeValue<'expr>>,
    dimension: usize,
) -> EvaluationResult<Result<(usize, MatchKind), ScalarError>>
where
    R: Resolver + ?Sized,
{
    if let Some(error) = formula_error_of(index).or_else(|| range.and_then(formula_error_of)) {
        return Ok(Err(error));
    }
    let mut known_index = None;
    if !requires_reference_value(index) {
        known_index = Some(match scalar_integer_borrowed(evaluator, index)? {
            Ok(value) => match table_index(value, dimension) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            },
            Err(error) => return Ok(Err(error)),
        });
    }
    let mut known_range = None;
    if let Some(range) = range {
        if !requires_reference_value(range) {
            known_range = Some(match range_lookup_borrowed(evaluator, range)? {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            });
        }
    } else {
        known_range = Some(MatchKind::Ascending);
    }
    let index = match known_index {
        Some(value) => value,
        None => match scalar_integer_borrowed(evaluator, index)? {
            Ok(value) => match table_index(value, dimension) {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            },
            Err(error) => return Ok(Err(error)),
        },
    };
    let range = match known_range {
        Some(value) => value,
        None => match range {
            Some(range) => match range_lookup_borrowed(evaluator, range)? {
                Ok(value) => value,
                Err(error) => return Ok(Err(error)),
            },
            None => MatchKind::Ascending,
        },
    };
    Ok(Ok((index, range)))
}

fn requires_reference_value(value: &RuntimeValue<'_>) -> bool {
    matches!(
        value,
        RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) | RuntimeValue::SourceReference
    )
}

fn table_index(index: usize, dimension: usize) -> Result<usize, ScalarError> {
    if index == 0 {
        Err(ScalarError::Value)
    } else if index > dimension {
        Err(ScalarError::Reference)
    } else {
        Ok(index)
    }
}

fn scalar_integer_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<Result<usize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = scalar_argument_borrowed(evaluator, value)?;
    if matches!(value, RuntimeValue::Missing) {
        return Ok(Err(ScalarError::Value));
    }
    match super::integer_value(evaluator, Some(value))? {
        Ok(Some(value)) if value >= 0.0 => Ok(match usize::try_from(value as u128) {
            Ok(value) => Ok(value),
            Err(_) => Err(ScalarError::Number),
        }),
        Ok(Some(_)) => Ok(Err(ScalarError::Value)),
        Ok(None) => Ok(Err(ScalarError::Value)),
        Err(error) => Ok(Err(error)),
    }
}

fn range_lookup_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<Result<MatchKind, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = scalar_argument_borrowed(evaluator, value)?;
    match super::optional_logical(evaluator, Some(value), true)? {
        Ok(value) => Ok(Ok(if value {
            MatchKind::Ascending
        } else {
            MatchKind::Exact
        })),
        Err(error) => Ok(Err(error)),
    }
}

fn match_type_borrowed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<Result<MatchKind, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = scalar_argument_borrowed(evaluator, value)?;
    if matches!(value, RuntimeValue::Missing) {
        return Ok(Ok(MatchKind::Ascending));
    }
    match super::integer_value(evaluator, Some(value))? {
        Ok(Some(value)) => Ok(match value as i32 {
            -1 => Ok(MatchKind::Descending),
            0 => Ok(MatchKind::Exact),
            1 => Ok(MatchKind::Ascending),
            _ => Err(ScalarError::Value),
        }),
        Ok(None) => Ok(Err(ScalarError::Value)),
        Err(error) => Ok(Err(error)),
    }
}

fn apply_table_lookup<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
    axis: Axis,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(3..=4).contains(&arguments.len()) {
        return invalid_arity(evaluator, node, arguments);
    }
    // Move the key after validation so computed owned text needs no second
    // allocation or storage reservation on the scalar ownership-taking path.
    let lookup = arguments.remove(0);
    let lookup_error = formula_error_of(&lookup);
    let data = arguments.remove(0);
    let grid = match grid(evaluator, &data)? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let dimension = match axis {
        Axis::Row => grid.shape.rows(),
        Axis::Column => grid.shape.columns(),
    };
    let (index, range) =
        match table_controls(evaluator, &arguments[0], arguments.get(1), dimension)? {
            Ok(controls) => controls,
            Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
        };
    if axis.length(grid.shape) == 0 {
        return Ok(formula_error(
            lookup_error.unwrap_or(ScalarError::NotAvailable),
        ));
    }
    let lookup = scalar_argument(evaluator, lookup)?;
    finish_table_lookup(evaluator, grid, axis, index, range, lookup)
}

fn apply_lookup<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(2..=3).contains(&arguments.len()) {
        return invalid_arity(evaluator, node, arguments);
    }
    let lookup = scalar_argument(evaluator, arguments.remove(0))?;
    let lookup_error = formula_error_of(&lookup);
    let searched = arguments.remove(0);
    let searched = match grid(evaluator, &searched)? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let search_axis = match searched.vector_axis() {
        Ok(axis) => axis,
        Err(_error) => {
            // Both LOOKUP spellings select the first column of a square/tall
            // table or the first row of a wider table. Only Results must be a
            // vector in the three-argument form.
            if searched.shape.rows() >= searched.shape.columns() {
                Axis::Column
            } else {
                Axis::Row
            }
        },
    };
    let result_grid = if let Some(results) = arguments.first()
        && !matches!(results, RuntimeValue::Missing)
    {
        let grid = match grid(evaluator, results)? {
            Ok(grid) => grid,
            Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
        };
        if grid.vector_axis().is_err() {
            return Ok(formula_error(lookup_error.unwrap_or(ScalarError::Value)));
        }
        Some(grid)
    } else {
        None
    };

    let find = match key_from_runtime(evaluator, &lookup)? {
        Ok(key) => key,
        Err(error) => return Ok(error.into_runtime()),
    };
    let matched = match find_match(evaluator, searched, search_axis, find, MatchKind::Ascending)? {
        Ok(result) => result,
        Err(error) => return Ok(error.into_runtime()),
    };

    if let Some(result_grid) = result_grid {
        let Some(result_axis) = result_grid.vector_axis().ok() else {
            return Ok(formula_error(ScalarError::Value));
        };
        let result_length = result_axis.length(result_grid.shape);
        let result_grid = if matched.index >= result_length {
            match extend_result_reference(
                evaluator,
                result_grid,
                matched.index,
                search_axis.length(searched.shape),
            )? {
                Ok(grid) => grid,
                Err(error) => return Ok(formula_error(error)),
            }
        } else {
            result_grid
        };
        let (row, column) = match result_coordinate(result_grid, matched.index) {
            Ok(coordinate) => coordinate,
            Err(error) => return Ok(formula_error(error)),
        };
        return selected_value(evaluator, result_grid, row, column);
    }

    let (row, column) = match search_axis {
        Axis::Column => (matched.index, searched.shape.columns() - 1),
        Axis::Row => (searched.shape.rows() - 1, matched.index),
    };
    selected_value(evaluator, searched, row, column)
}

fn apply_match<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if !(2..=3).contains(&arguments.len()) {
        return invalid_arity(evaluator, node, arguments);
    }
    // Validate MATCH's mode before reading a reference lookup key.
    let lookup_raw = arguments.remove(0);
    let lookup_error = formula_error_of(&lookup_raw);
    let search = arguments.remove(0);
    let match_kind = if let Some(value) = arguments.pop() {
        match match_type(evaluator, value)? {
            Ok(kind) => kind,
            Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
        }
    } else {
        MatchKind::Ascending
    };
    let grid = match grid(evaluator, &search)? {
        Ok(grid) => grid,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let axis = match grid.vector_axis() {
        Ok(axis) => axis,
        Err(error) => return Ok(formula_error(lookup_error.unwrap_or(error))),
    };
    let lookup = scalar_argument(evaluator, lookup_raw)?;
    let find = match key_from_runtime(evaluator, &lookup)? {
        Ok(key) => key,
        Err(error) => return Ok(error.into_runtime()),
    };
    let matched = match find_match(evaluator, grid, axis, find, match_kind)? {
        Ok(result) => result,
        Err(error) => return Ok(error.into_runtime()),
    };
    Ok(RuntimeValue::Scalar(WorkingValue::Number(
        matched.index as f64 + 1.0,
    )))
}

fn invalid_arity<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if let Some(error) = first_formula_error(evaluator, &arguments)? {
        return Ok(formula_error(error));
    }
    if node.child_count() != 0 {
        evaluator.scalar.charge_work(node.child_count() as u64)?;
    }
    Ok(formula_error(ScalarError::Value))
}

fn first_formula_error<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<Option<ScalarError>>
where
    R: Resolver + ?Sized,
{
    for argument in arguments {
        match argument {
            RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(Some(*error)),
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
            | RuntimeValue::Areas(_)
            | RuntimeValue::SourceReference => {},
        }
    }
    Ok(None)
}

fn formula_error_of(value: &RuntimeValue<'_>) -> Option<ScalarError> {
    match value {
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Some(*error),
        _ => None,
    }
}

fn grid<'value, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &'value RuntimeValue<'expr>,
) -> EvaluationResult<Result<SearchGrid<'value, 'expr>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let result = match value {
        RuntimeValue::Array(array) => SearchGrid {
            source: SearchSource::Array(array),
            shape: array.shape,
            reference: None,
        },
        RuntimeValue::Areas(areas) => {
            if areas.is_list || areas.areas.len() != 1 {
                return Ok(Err(ScalarError::Value));
            }
            let Some(area) = areas.areas.first().copied() else {
                return Ok(Err(ScalarError::Value));
            };
            SearchGrid {
                source: SearchSource::Area(area),
                shape: Shape::new(area.rect.rows(), area.rect.columns())?,
                reference: Some(area),
            }
        },
        RuntimeValue::ScalarCell(area) => SearchGrid {
            source: SearchSource::Area(*area),
            shape: Shape::new(1, 1)?,
            reference: Some(*area),
        },
        RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(Err(*error)),
        RuntimeValue::Empty | RuntimeValue::Missing | RuntimeValue::Scalar(_) => {
            return Ok(Err(ScalarError::Value));
        },
        RuntimeValue::SourceReference => {
            return Err(EvaluationFailure::Unsupported(
                super::super::super::UnsupportedKind::Reference,
            ));
        },
    };
    let _ = evaluator;
    Ok(Ok(result))
}

fn extend_result_reference<'value, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'value, 'expr>,
    selected_index: usize,
    search_length: usize,
) -> EvaluationResult<Result<SearchGrid<'value, 'expr>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(area) = grid.reference else {
        return Ok(Ok(grid));
    };
    let rows = area.rect.rows();
    let columns = area.rect.columns();
    if rows > 1 && columns > 1 {
        return Ok(Err(ScalarError::Value));
    }
    let length = rows.max(columns);
    let target = search_length;
    if selected_index < length {
        return Ok(Ok(grid));
    }
    if length >= target {
        return Ok(Ok(grid));
    }
    let extend_column = rows == 1 && columns == 1;
    let (new_rows, new_columns) = if extend_column {
        (target, 1)
    } else if rows == 1 {
        (1, target)
    } else {
        (target, 1)
    };
    let new_row_end = area
        .rect
        .row_start
        .checked_add(new_rows)
        .ok_or(ScalarError::NotAvailable);
    let Ok(new_row_end) = new_row_end else {
        return Ok(Err(ScalarError::NotAvailable));
    };
    let new_column_end = area
        .rect
        .column_start
        .checked_add(new_columns)
        .ok_or(ScalarError::NotAvailable);
    let Ok(new_column_end) = new_column_end else {
        return Ok(Err(ScalarError::NotAvailable));
    };
    let cells = new_rows
        .checked_mul(new_columns)
        .ok_or(ScalarError::NotAvailable);
    let Ok(cells) = cells else {
        return Ok(Err(ScalarError::NotAvailable));
    };
    // Reject impossible geometry before asking the resolver for sheet
    // metadata.  This is the same admission order used by ordinary retained
    // references and keeps a limit refusal independent of provider state.
    evaluator.check_reference_cells(cells)?;
    evaluator.check_reference_areas(1)?;
    let sheet_name = match area.sheet {
        SheetRef::Current => evaluator.position.sheet,
        SheetRef::Named(name) => name,
    };
    let extent = resolver_sheet_extent(evaluator, sheet_name)?;
    let Some(extent) = extent else {
        return Ok(Err(ScalarError::NotAvailable));
    };
    if new_row_end > extent.rows() || new_column_end > extent.columns() {
        return Ok(Err(ScalarError::NotAvailable));
    }
    let area = RuntimeArea {
        sheet: area.sheet,
        sheet_index: area.sheet_index,
        rect: super::super::Rect::new(
            area.rect.row_start,
            new_row_end,
            area.rect.column_start,
            new_column_end,
        )?,
    };
    Ok(Ok(SearchGrid {
        source: SearchSource::Area(area),
        shape: Shape::new(new_rows, new_columns)?,
        reference: Some(area),
    }))
}

/// Resolve extension geometry through the same bounded metadata seam as the
/// other reference operators.  A successful provider call gets a post-call
/// cancellation fence; a typed provider failure is returned unchanged.
fn resolver_sheet_extent<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
) -> EvaluationResult<Option<SheetExtent>>
where
    R: Resolver + ?Sized,
{
    evaluator.scalar.charge_work(1)?;
    let result = evaluator.resolver.sheet_extent(name, evaluator.execution);
    match result {
        Ok(value) => {
            evaluator.scalar.charge_work(0)?;
            Ok(value)
        },
        Err(error) => Err(error),
    }
}

fn result_coordinate(
    grid: SearchGrid<'_, '_>,
    index: usize,
) -> Result<(usize, usize), ScalarError> {
    let axis = if grid.shape.rows() == 1 {
        Axis::Row
    } else if grid.shape.columns() == 1 {
        Axis::Column
    } else {
        return Err(ScalarError::Value);
    };
    if index >= axis.length(grid.shape) {
        return Err(ScalarError::NotAvailable);
    }
    Ok(axis.coordinate(index))
}

fn scalar_argument<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    evaluator.matrix_scalar_parameter(value)
}

fn match_type<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<Result<MatchKind, ScalarError>>
where
    R: Resolver + ?Sized,
{
    if matches!(value, RuntimeValue::Missing) {
        return Ok(Ok(MatchKind::Ascending));
    }
    match super::integer_value(evaluator, Some(value))? {
        Ok(Some(value)) => Ok(match value as i32 {
            -1 => Ok(MatchKind::Descending),
            0 => Ok(MatchKind::Exact),
            1 => Ok(MatchKind::Ascending),
            _ => Err(ScalarError::Value),
        }),
        Ok(None) => Ok(Err(ScalarError::Value)),
        Err(error) => Ok(Err(error)),
    }
}

fn key_from_runtime<'value, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &'value RuntimeValue<'expr>,
) -> EvaluationResult<Result<scalar_lookup::SearchValue<'value>, SearchError>>
where
    R: Resolver + ?Sized,
{
    Ok(match value {
        RuntimeValue::Empty => Ok(scalar_lookup::SearchValue::Number(0.0)),
        RuntimeValue::Missing => Err(SearchError::Generated(ScalarError::Value)),
        RuntimeValue::Scalar(WorkingValue::Number(value)) if value.is_finite() => {
            Ok(scalar_lookup::SearchValue::Number(*value))
        },
        RuntimeValue::Scalar(WorkingValue::Number(_)) => {
            Err(SearchError::Generated(ScalarError::Number))
        },
        RuntimeValue::Scalar(WorkingValue::Logical(value)) => {
            Ok(scalar_lookup::SearchValue::Logical(*value))
        },
        RuntimeValue::Scalar(WorkingValue::Text(value)) => {
            let length = value.text.len();
            evaluator.scalar.charge_bytes(length)?;
            Ok(scalar_lookup::SearchValue::Text(value.text.as_ref()))
        },
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Err(SearchError::Formula(*error)),
        RuntimeValue::Scalar(WorkingValue::Complex(_)) => {
            Err(SearchError::Generated(ScalarError::Value))
        },
        RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
            Err(SearchError::Generated(ScalarError::Value))
        },
        RuntimeValue::SourceReference => {
            return Err(EvaluationFailure::Unsupported(
                super::super::super::UnsupportedKind::Reference,
            ));
        },
    })
}

fn cell_key<'a>(cell: SearchCell<'a>) -> Result<scalar_lookup::SearchValue<'a>, SearchError> {
    match cell {
        SearchCell::Empty => Ok(scalar_lookup::SearchValue::Number(0.0)),
        SearchCell::Missing => Err(SearchError::Generated(ScalarError::NotAvailable)),
        SearchCell::Number(value) if value.is_finite() => {
            Ok(scalar_lookup::SearchValue::Number(value))
        },
        SearchCell::Number(_) => Err(SearchError::Generated(ScalarError::Number)),
        SearchCell::Logical(value) => Ok(scalar_lookup::SearchValue::Logical(value)),
        SearchCell::Text(value) => Ok(scalar_lookup::SearchValue::Text(value)),
        SearchCell::Formula(error) => Err(SearchError::Formula(error)),
        SearchCell::Generated(error) => Err(SearchError::Generated(error)),
        SearchCell::Complex => Err(SearchError::Generated(ScalarError::Value)),
    }
}

fn find_match<'key, 'grid, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'grid, 'expr>,
    axis: Axis,
    find: scalar_lookup::SearchValue<'key>,
    kind: MatchKind,
) -> EvaluationResult<Result<MatchResult, SearchError>>
where
    R: Resolver + ?Sized,
    'expr: 'grid,
{
    let len = axis.length(grid.shape);
    match kind {
        MatchKind::Exact => {
            let mut first_error = None;
            for index in 0..len {
                let candidate = read_key(evaluator, grid, axis, index)?;
                let candidate = match candidate {
                    Ok(candidate) => candidate,
                    Err(error) => {
                        first_error.get_or_insert(error);
                        continue;
                    },
                };
                let equal = match scalar_lookup::exact_match(candidate, find, |units| {
                    evaluator.scalar.charge_work(units)
                })? {
                    Ok(equal) => equal,
                    Err(error) => {
                        first_error.get_or_insert(SearchError::Generated(error));
                        continue;
                    },
                };
                if equal && first_error.is_none() {
                    return Ok(Ok(MatchResult { index }));
                }
            }
            if let Some(error) = first_error {
                return Ok(Err(error));
            }
            Ok(Err(SearchError::Generated(ScalarError::NotAvailable)))
        },
        MatchKind::Ascending | MatchKind::Descending => {
            let ascending = kind == MatchKind::Ascending;
            let mut low = 0usize;
            let mut high = len;
            let mut best = None;
            let mut best_candidate = None;
            while low < high {
                let middle = low + (high - low) / 2;
                let candidate = read_key(evaluator, grid, axis, middle)?;
                let candidate = match candidate {
                    Ok(candidate) => candidate,
                    Err(error) => return Ok(Err(error)),
                };
                let ordering = match scalar_lookup::compare_values(candidate, find, |units| {
                    evaluator.scalar.charge_work(units)
                })? {
                    Ok(ordering) => ordering,
                    Err(error) => return Ok(Err(SearchError::Generated(error))),
                };
                let qualifies = if ascending {
                    ordering != std::cmp::Ordering::Greater
                } else {
                    ordering != std::cmp::Ordering::Less
                };
                if qualifies {
                    best = Some(middle);
                    best_candidate = Some(candidate);
                    low = middle.saturating_add(1);
                } else {
                    high = middle;
                }
            }
            let Some(index) = best else {
                return Ok(Err(SearchError::Generated(ScalarError::NotAvailable)));
            };
            let Some(candidate) = best_candidate else {
                return Err(EvaluationFailure::InvalidExpression(
                    "lookup binary search lost its selected candidate",
                ));
            };
            if !scalar_lookup::approximate_type_compatible(find, candidate, !ascending) {
                return Ok(Err(SearchError::Generated(ScalarError::NotAvailable)));
            }
            Ok(Ok(MatchResult { index }))
        },
    }
}

fn read_key<'grid, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'grid, 'expr>,
    axis: Axis,
    index: usize,
) -> EvaluationResult<Result<scalar_lookup::SearchValue<'grid>, SearchError>>
where
    R: Resolver + ?Sized,
    'expr: 'grid,
{
    let (row, column) = axis.coordinate(index);
    let cell = read_cell(evaluator, grid, row, column)?;
    Ok(match cell {
        Ok(cell) => cell_key(cell),
        Err(error) => Err(error),
    })
}

fn read_cell<'grid, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'grid, 'expr>,
    row: usize,
    column: usize,
) -> EvaluationResult<Result<SearchCell<'grid>, SearchError>>
where
    R: Resolver + ?Sized,
    'expr: 'grid,
{
    if row >= grid.shape.rows() || column >= grid.shape.columns() {
        return Ok(Err(SearchError::Generated(ScalarError::NotAvailable)));
    }
    let index = row
        .checked_mul(grid.shape.columns())
        .and_then(|offset| offset.checked_add(column))
        .ok_or(EvaluationFailure::InvalidExpression(
            "lookup cell index overflows",
        ))?;
    evaluator.charge_cell_work(index)?;
    match grid.source {
        SearchSource::Array(array) => {
            Ok(Ok(search_cell_from_element(array.cells.get(index).ok_or(
                EvaluationFailure::InvalidExpression("lookup array cell is missing"),
            )?)))
        },
        SearchSource::Area(area) => {
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start + row,
                area.rect.column_start + column,
            )?;
            Ok(Ok(search_cell_from_read(evaluator, read)?))
        },
    }
}

fn search_cell_from_element<'cell, 'expr>(
    element: &'cell RuntimeElement<'expr>,
) -> SearchCell<'cell> {
    match element {
        RuntimeElement::Empty => SearchCell::Empty,
        RuntimeElement::Missing => SearchCell::Missing,
        RuntimeElement::Present(WorkingValue::Number(value)) => SearchCell::Number(*value),
        RuntimeElement::Present(WorkingValue::Logical(value)) => SearchCell::Logical(*value),
        RuntimeElement::Present(WorkingValue::Text(value)) => SearchCell::Text(value.text.as_ref()),
        RuntimeElement::Present(WorkingValue::Error(error)) => SearchCell::Formula(*error),
        RuntimeElement::Present(WorkingValue::Complex(_)) => SearchCell::Complex,
    }
}

fn search_cell_from_read<'cell, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    read: CellRead<'expr>,
) -> EvaluationResult<SearchCell<'cell>>
where
    R: Resolver + ?Sized,
    'expr: 'cell,
{
    // `read_to_element` owns the common finite-number, text-budget, and
    // borrowed-text conversion path.  Keep the source text from `read` so the
    // search state borrows the resolver value instead of borrowing a local
    // RuntimeElement temporary.  Non-finite numbers need the search-specific
    // generated #NUM distinction, which the generic element conversion
    // intentionally represents as a formula error.
    let nonfinite_number = matches!(read, CellRead::Number(value) if !value.is_finite());
    let text = match read {
        CellRead::Text(text) => Some(text),
        CellRead::Empty
        | CellRead::Number(_)
        | CellRead::Logical(_)
        | CellRead::Error(_)
        | CellRead::Unsupported => None,
    };
    let element = evaluator.read_to_element(read)?;
    if nonfinite_number {
        return Ok(SearchCell::Generated(ScalarError::Number));
    }
    if let Some(text) = text {
        return Ok(SearchCell::Text(text));
    }
    Ok(match element {
        RuntimeElement::Empty => SearchCell::Empty,
        RuntimeElement::Missing => SearchCell::Missing,
        RuntimeElement::Present(WorkingValue::Number(value)) => SearchCell::Number(value),
        RuntimeElement::Present(WorkingValue::Logical(value)) => SearchCell::Logical(value),
        RuntimeElement::Present(WorkingValue::Text(_)) => {
            return Err(EvaluationFailure::InvalidExpression(
                "lookup resolver text lost its borrowed source",
            ));
        },
        RuntimeElement::Present(WorkingValue::Error(error)) => SearchCell::Formula(error),
        RuntimeElement::Present(WorkingValue::Complex(_)) => SearchCell::Complex,
    })
}

fn selected_value<'value, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    grid: SearchGrid<'value, 'expr>,
    row: usize,
    column: usize,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if row >= grid.shape.rows() || column >= grid.shape.columns() {
        return Ok(formula_error(ScalarError::NotAvailable));
    }
    let index = row
        .checked_mul(grid.shape.columns())
        .and_then(|offset| offset.checked_add(column))
        .ok_or(EvaluationFailure::InvalidExpression(
            "lookup selected index overflows",
        ))?;
    match grid.source {
        SearchSource::Array(array) => {
            evaluator.charge_cell_work(index)?;
            let element = array
                .cells
                .get(index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "lookup selected array cell is missing",
                ))?;
            Ok(super::super::element_to_runtime(
                evaluator.clone_element(element)?,
            ))
        },
        SearchSource::Area(area) => {
            evaluator.charge_cell_work(index)?;
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start + row,
                area.rect.column_start + column,
            )?;
            // Keep the selected-cell read explicit.  A search candidate and
            // its returned result may alias, but the result still crosses the
            // normal read_to_element conversion and provider/cancellation
            // fence exactly once after the search has selected it.
            let element = evaluator.read_to_element(read)?;
            Ok(super::super::element_to_runtime(element))
        },
    }
}
