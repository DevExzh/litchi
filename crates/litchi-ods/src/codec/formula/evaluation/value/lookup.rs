//! Resolver-aware OpenFormula lookup functions.
//!
//! Lookup data sources remain either scalar values, materialized arrays, or
//! first-class reference descriptors.  This module owns the value-side
//! distinction.  In particular, INDEX and OFFSET construct descriptors and
//! never read their source cells; the search functions delegate ordered
//! traversal to [`search`], which reads only the cells needed by the search
//! and selected result.

mod indirect;
mod lifting;
mod search;

use super::super::lookup::{self, Function};
use super::{
    EvaluationFailure, EvaluationResult, Rect, ReferenceOperandKind, Resolver, RuntimeArea,
    RuntimeAreaSet, RuntimeArrayValue, RuntimeElement, RuntimeValue, ScalarError, Shape,
    ValueEvaluator, WorkingValue, ensure_capacity,
};

/// Return whether `name` belongs to the nine-function lookup family.
pub(super) fn is_lookup_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

/// Dynamic reference constructors need their result geometry before a
/// projected matrix branch can publish its output shape.  Keep these on the
/// dedicated probe path; a generic child walk either guesses one cell or
/// broadcasts selector geometry into the descriptor result.
pub(super) fn is_dynamic_shape_function(name: &str) -> bool {
    matches!(
        Function::from_name(name),
        Some(Function::Index | Function::Offset | Function::Indirect)
    )
}

/// Whether a dynamic lookup's scalar arguments can be reused when the outer
/// projected demand widens.  A direct scalar/reference argument depends on
/// the absolute caller position, which is part of the probe key; an array or
/// computed expression may map differently under a new demand shape and is
/// therefore kept demand-sensitive.
pub(super) fn probe_demand_invariant(node: super::super::Node<'_>) -> bool {
    let Some(function) = node.function_name().and_then(Function::from_name) else {
        return false;
    };
    // Keep this syntax preflight shallow. A parenthesized/computed subtree is
    // conservatively demand-sensitive; the runtime probe still handles it
    // correctly, but it is not reused across shape refinement.
    let scalar = |child: super::super::Node<'_>| {
        matches!(
            child.kind(),
            super::super::Kind::Number
                | super::super::Kind::String
                | super::super::Kind::Error
                | super::super::Kind::Missing
                | super::super::Kind::Reference(_)
        )
    };
    match function {
        Function::Indirect => {
            (0..node.child_count()).all(|index| node.child(index).is_some_and(scalar))
        },
        Function::Index => {
            node.child(0).is_some_and(scalar)
                && (1..node.child_count()).all(|index| node.child(index).is_some_and(scalar))
        },
        Function::Offset => {
            node.child(0).is_some_and(scalar)
                && (1..node.child_count()).all(|index| node.child(index).is_some_and(scalar))
        },
        _ => false,
    }
}

pub(super) fn may_return_reference(name: &str) -> bool {
    matches!(
        Function::from_name(name),
        Some(Function::Index | Function::Offset | Function::Indirect | Function::Choose)
    )
}

/// Keep data/reference arguments complete in a projected matrix branch.  The
/// scalar selector arguments still use the enclosing position and are
/// consumed through the ordinary value path.
pub(super) fn matrix_argument(name: &str, index: usize, _node: super::super::Node<'_>) -> bool {
    let Some(function) = Function::from_name(name) else {
        return false;
    };
    matches!(
        function,
        Function::HLookup | Function::Lookup | Function::Match | Function::VLookup
    ) && function.matrix_argument(index)
}

/// Keep descriptor inputs complete while retaining the caller's scalar or
/// matrix mode. INDEX and OFFSET have Reference/Any signatures rather than a
/// ForceArray signature, so their computed arguments must not be forced into
/// a new matrix context.
pub(super) fn complete_argument(name: &str, index: usize) -> bool {
    Function::from_name(name).is_some_and(|function| {
        function.reference_argument(index)
            && !matches!(
                function,
                Function::HLookup | Function::Lookup | Function::Match | Function::VLookup
            )
    })
}

/// Reference-returning functions consume one scalar parameter instead of
/// iterating that parameter at the surrounding projected output coordinate.
pub(super) fn scalar_argument(name: &str, index: usize) -> bool {
    Function::from_name(name).is_some_and(|function| match function {
        Function::Index | Function::Offset => index > 0,
        Function::Indirect => true,
        _ => false,
    })
}

/// In a projected lazy branch, evaluate a computed lookup key/selector in
/// scalar context so nested reducers see the requested worksheet position.
/// An inline Array remains in the caller's projection: `visit_array` then
/// selects its coordinate and the lookup mapper can preserve ordinary array
/// broadcasting.  This split is needed for `MATCH(SUM(MUNIT(...)); ...)` and
/// the equivalent selector/mode positions without collapsing literal arrays
/// to their first element.
pub(super) fn projected_scalar_argument<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    index: usize,
    node: super::super::Node<'expr>,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    let Some(function) = Function::from_name(name) else {
        return Ok(false);
    };
    let scalar = match function {
        Function::HLookup | Function::VLookup => matches!(index, 0 | 2 | 3),
        Function::Lookup => index == 0,
        Function::Match => matches!(index, 0 | 2),
        _ => false,
    };
    if !scalar {
        return Ok(false);
    }
    projected_scalar_node(evaluator, node)
}

#[derive(Clone, Copy)]
enum ScalarShapeFrame<'expr> {
    Enter(super::super::Node<'expr>),
    Reduce(usize),
}

/// Classify a lookup scalar expression without evaluating it.
///
/// The walk is deliberately iterative and charges one unit per AST node. A
/// reducer is a scalar barrier: its complete sequence input does not widen
/// the lookup key shape. Ordinary operators and scalar functions recurse so a
/// wrapper such as `ABS(SUM(MUNIT(...)))` inherits the reducer's scalar shape.
/// Explicit array-producing matrix functions and literal arrays remain
/// matrix-shaped, preserving per-coordinate lookup keys.
fn projected_scalar_node<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    root: super::super::Node<'expr>,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    let maximum = evaluator.limits.scalar.max_stack_entries;
    // Declare reservations before their vectors so vectors drop first on all
    // exits, including resource and cancellation failures.
    let mut frame_reservation = None;
    let mut frames = Vec::new();
    let mut result_reservation = None;
    let mut results = Vec::new();
    ensure_capacity(
        &mut frames,
        &mut frame_reservation,
        1,
        maximum,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula lookup scalar-shape frames",
    )?;
    frames.push(ScalarShapeFrame::Enter(root));

    while let Some(frame) = frames.pop() {
        match frame {
            ScalarShapeFrame::Enter(node) => {
                evaluator.scalar.charge_work(1)?;
                let children = match node.kind() {
                    // These values are scalar at the lookup argument boundary.
                    super::super::Kind::Number
                    | super::super::Kind::String
                    | super::super::Kind::Error
                    | super::super::Kind::Missing
                    | super::super::Kind::Reference(_) => {
                        ensure_capacity(
                            &mut results,
                            &mut result_reservation,
                            1,
                            maximum,
                            evaluator.execution,
                            &evaluator.storage_budget,
                            "formula lookup scalar-shape results",
                        )?;
                        results.push(true);
                        continue;
                    },
                    // A reference operator produces a reference sequence, but
                    // scalar lookup conversion applies implicit intersection
                    // at the requested projected position.
                    super::super::Kind::Infix(
                        super::super::InfixOperator::Range
                        | super::super::InfixOperator::Intersection
                        | super::super::InfixOperator::Union,
                    ) => {
                        ensure_capacity(
                            &mut results,
                            &mut result_reservation,
                            1,
                            maximum,
                            evaluator.execution,
                            &evaluator.storage_budget,
                            "formula lookup scalar-shape results",
                        )?;
                        results.push(true);
                        continue;
                    },
                    // Inline arrays are the primary matrix-key spelling. They
                    // must retain the caller's projection for broadcasting.
                    super::super::Kind::Array(_)
                    | super::super::Kind::ArrayRow
                    | super::super::Kind::QuotedLabel
                    | super::super::Kind::AutomaticIntersection
                    | super::super::Kind::NamedExpression { .. } => {
                        ensure_capacity(
                            &mut results,
                            &mut result_reservation,
                            1,
                            maximum,
                            evaluator.execution,
                            &evaluator.storage_budget,
                            "formula lookup scalar-shape results",
                        )?;
                        results.push(false);
                        continue;
                    },
                    super::super::Kind::Parenthesized
                    | super::super::Kind::Prefix(_)
                    | super::super::Kind::Postfix(_) => 1,
                    super::super::Kind::Infix(_) => 2,
                    super::super::Kind::Function { name } => {
                        if scalar_result_barrier(name) {
                            ensure_capacity(
                                &mut results,
                                &mut result_reservation,
                                1,
                                maximum,
                                evaluator.execution,
                                &evaluator.storage_budget,
                                "formula lookup scalar-shape results",
                            )?;
                            results.push(true);
                            continue;
                        }
                        if super::matrix::is_matrix_function(name) {
                            ensure_capacity(
                                &mut results,
                                &mut result_reservation,
                                1,
                                maximum,
                                evaluator.execution,
                                &evaluator.storage_budget,
                                "formula lookup scalar-shape results",
                            )?;
                            results.push(false);
                            continue;
                        }
                        node.child_count()
                    },
                };
                if children == 0 {
                    ensure_capacity(
                        &mut results,
                        &mut result_reservation,
                        1,
                        maximum,
                        evaluator.execution,
                        &evaluator.storage_budget,
                        "formula lookup scalar-shape results",
                    )?;
                    results.push(true);
                    continue;
                }
                let frame_count =
                    children
                        .checked_add(1)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "lookup scalar-shape frame overflow",
                        ))?;
                ensure_capacity(
                    &mut frames,
                    &mut frame_reservation,
                    frame_count,
                    maximum,
                    evaluator.execution,
                    &evaluator.storage_budget,
                    "formula lookup scalar-shape frames",
                )?;
                frames.push(ScalarShapeFrame::Reduce(children));
                for index in (0..children).rev() {
                    let child = node
                        .child(index)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "lookup scalar-shape child is missing",
                        ))?;
                    frames.push(ScalarShapeFrame::Enter(child));
                }
            },
            ScalarShapeFrame::Reduce(children) => {
                if results.len() < children {
                    return Err(EvaluationFailure::InvalidExpression(
                        "lookup scalar-shape result is missing",
                    ));
                }
                let start = results.len() - children;
                let scalar = results[start..].iter().all(|value| *value);
                results.truncate(start);
                ensure_capacity(
                    &mut results,
                    &mut result_reservation,
                    1,
                    maximum,
                    evaluator.execution,
                    &evaluator.storage_budget,
                    "formula lookup scalar-shape results",
                )?;
                results.push(scalar);
            },
        }
    }
    let result = results.pop().ok_or(EvaluationFailure::InvalidExpression(
        "lookup scalar-shape result is missing",
    ))?;
    drop(results);
    drop(result_reservation);
    drop(frames);
    drop(frame_reservation);
    Ok(result)
}

/// Classify an arbitrary scalar expression for a projected lookup/date
/// parameter.  The walk is shared with lookup selectors so date parameters
/// inherit the same position-sensitive treatment for nested `MUNIT` and
/// explicit arrays without duplicating the bounded AST probe.
pub(super) fn projected_scalar_expression<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    projected_scalar_node(evaluator, node)
}

fn scalar_result_barrier(name: &str) -> bool {
    super::aggregate::is_aggregate_function(name)
        || super::statistical::is_statistical_function(name)
        || super::descriptive::is_descriptive_function(name)
        || super::paired::is_paired_function(name)
        || super::conditional::is_conditional_function(name)
        || super::database::is_database_function(name)
        || super::complex::is_complex_sequence_function(name)
        || super::super::inspection::Function::from_name(name).is_some_and(|function| {
            matches!(
                function,
                super::super::inspection::Function::N
                    | super::super::inspection::Function::Na
                    | super::super::inspection::Function::Type
            )
        })
        || name.eq_ignore_ascii_case("MDETERM")
}

/// A nested lookup's data vectors are complete reducer inputs.  Computed
/// selectors remain position-sensitive, so they deliberately do not inherit
/// this flag in the demand-cache classifier.
pub(super) fn criterion_full_argument(
    name: &str,
    index: usize,
    _node: super::super::Node<'_>,
) -> bool {
    Function::from_name(name).is_some_and(|function| function.reference_argument(index))
}

/// Lookup results are conservative demand-cache candidates.  Descriptor
/// selection and all searches depend on argument geometry or read ordering;
/// ADDRESS is pure only when its scalar arguments have already been proven
/// invariant by the caller.  Returning false keeps position-sensitive nested
/// selectors from being reused across projected coordinates.
pub(super) fn cacheable_branch<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    node: super::super::Node<'expr>,
) -> EvaluationResult<bool> {
    let Some(function) = node.function_name().and_then(Function::from_name) else {
        return Ok(false);
    };
    if !matches!(
        function,
        Function::HLookup | Function::VLookup | Function::Lookup | Function::Match
    ) || !function.valid_arity(node.child_count())
    {
        return Ok(false);
    }
    for index in 0..node.child_count() {
        let Some(mut child) = node.child(index) else {
            return Ok(false);
        };
        loop {
            evaluator.scalar.charge_work(1)?;
            if !matches!(child.kind(), super::super::Kind::Parenthesized) {
                break;
            }
            let Some(inner) = child.child(0) else {
                return Ok(false);
            };
            child = inner;
        }
        let invariant = match child.kind() {
            super::super::Kind::Reference(reference) => {
                if function.reference_argument(index) {
                    ValueEvaluator::<R>::cacheable_reference(reference)
                } else {
                    ValueEvaluator::<R>::cacheable_condition_reference(reference)
                }
            },
            super::super::Kind::Number
            | super::super::Kind::String
            | super::super::Kind::Error
            | super::super::Kind::Missing => true,
            super::super::Kind::Function { name } => {
                child.child_count() == 0
                    && (name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE"))
            },
            // Computed selectors and Array keys may vary with the projected
            // coordinate, even beneath a complete reducer such as SUM(MUNIT()).
            _ => false,
        };
        if !invariant {
            return Ok(false);
        }
    }
    Ok(true)
}

pub(super) fn reference_kind(
    function: Function,
    data: Option<ReferenceOperandKind>,
) -> ReferenceOperandKind {
    match function {
        Function::Index | Function::Offset => match data {
            Some(ReferenceOperandKind::Reference) => ReferenceOperandKind::Reference,
            Some(ReferenceOperandKind::ReferenceList) if function == Function::Index => {
                ReferenceOperandKind::ReferenceList
            },
            Some(ReferenceOperandKind::ReferenceList) => {
                ReferenceOperandKind::Error(ScalarError::Value)
            },
            Some(ReferenceOperandKind::SourceReference) => ReferenceOperandKind::SourceReference,
            Some(ReferenceOperandKind::Array) => ReferenceOperandKind::Array,
            Some(ReferenceOperandKind::Error(error)) => ReferenceOperandKind::Error(error),
            Some(ReferenceOperandKind::Scalar) => ReferenceOperandKind::Scalar,
            _ => ReferenceOperandKind::Unknown,
        },
        Function::Indirect => ReferenceOperandKind::Reference,
        Function::Choose => ReferenceOperandKind::Unknown,
        Function::Address
        | Function::HLookup
        | Function::Lookup
        | Function::Match
        | Function::VLookup => ReferenceOperandKind::Scalar,
    }
}

pub(super) fn shape_base<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    function: Function,
) -> EvaluationResult<Option<Shape>>
where
    R: Resolver + ?Sized,
{
    match function {
        // ADDRESS and the search functions can iterate over scalar selector
        // arrays.  Their output shape is therefore supplied by the generic
        // lookup shape walk in the value VM; returning 1x1 here would erase
        // that shape before the walk sees the selectors.
        Function::Address
        | Function::HLookup
        | Function::Lookup
        | Function::Match
        | Function::VLookup => Ok(None),
        Function::Index => {
            let Some(argument) = node.child(0) else {
                return Ok(Some(Shape::new(1, 1)?));
            };
            let mut argument = argument;
            while matches!(argument.kind(), super::super::Kind::Parenthesized) {
                argument = argument
                    .child(0)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "INDEX shape parentheses are empty",
                    ))?;
            }
            if let super::super::Kind::Reference(reference) = argument.kind() {
                let Some(shape) = evaluator.reference_shape_hint(reference)? else {
                    return Ok(None);
                };
                return indexed_result_shape(node, shape);
            }
            Ok(None)
        },
        Function::Offset => {
            let Some(argument) = node.child(0) else {
                return Ok(Some(Shape::new(1, 1)?));
            };
            let mut argument = argument;
            while matches!(argument.kind(), super::super::Kind::Parenthesized) {
                argument = argument
                    .child(0)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "OFFSET shape parentheses are empty",
                    ))?;
            }
            if let super::super::Kind::Reference(reference) = argument.kind() {
                let Some(shape) = evaluator.reference_shape_hint(reference)? else {
                    return Ok(None);
                };
                return offset_result_shape(node, shape);
            }
            Ok(None)
        },
        Function::Indirect | Function::Choose => Ok(None),
    }
}

/// Determine the shape of a direct INDEX reference when all selectors needed
/// for the geometry are literal.  Selector conversion errors still publish a
/// scalar formula error, so an out-of-bounds literal is represented by 1x1.
/// Unknown/computed selectors return `None` and are left to the ordinary
/// runtime path rather than being evaluated speculatively by the planner.
fn indexed_result_shape<'expr>(
    node: super::super::Node<'expr>,
    source: Shape,
) -> EvaluationResult<Option<Shape>> {
    let Some(row) = literal_selector(node.child(1)) else {
        return Ok(None);
    };
    let Some(column) = literal_selector(node.child(2)) else {
        return Ok(None);
    };
    if row > source.rows() || column > source.columns() {
        return Ok(Some(Shape::new(1, 1)?));
    }
    let rows = if row == 0 { source.rows() } else { 1 };
    let columns = if column == 0 { source.columns() } else { 1 };
    Ok(Some(Shape::new(rows, columns)?))
}

fn offset_result_shape<'expr>(
    node: super::super::Node<'expr>,
    source: Shape,
) -> EvaluationResult<Option<Shape>> {
    let Some(height) = literal_extent(node.child(3), source.rows()) else {
        return Ok(None);
    };
    let Some(width) = literal_extent(node.child(4), source.columns()) else {
        return Ok(None);
    };
    Ok(Some(Shape::new(height, width)?))
}

fn literal_selector(node: Option<super::super::Node<'_>>) -> Option<usize> {
    let Some(node) = node else {
        return Some(0);
    };
    if node.is_missing() {
        return Some(0);
    }
    let super::super::Kind::Number = node.kind() else {
        return None;
    };
    let value = node.text().parse::<f64>().ok()?;
    if !value.is_finite() {
        return Some(1);
    }
    let value = value.trunc();
    if value == 0.0 {
        return Some(0);
    }
    if value < 0.0 || value >= (usize::MAX as u128 + 1) as f64 {
        return None;
    }
    Some(value as usize)
}

fn literal_extent(node: Option<super::super::Node<'_>>, default: usize) -> Option<usize> {
    let Some(node) = node else {
        return Some(default);
    };
    if node.is_missing() {
        return Some(default);
    }
    let super::super::Kind::Number = node.kind() else {
        return None;
    };
    let value = node.text().parse::<f64>().ok()?;
    if !value.is_finite()
        || value.trunc() <= 0.0
        || value.trunc() >= (usize::MAX as u128 + 1) as f64
    {
        return None;
    }
    Some(value.trunc() as usize)
}

/// Apply one value-side lookup function after its scheduled arguments have
/// been visited.  Formula errors are values; resolver, cancellation, source,
/// and resource failures remain typed evaluation failures.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let function = Function::from_name(name).ok_or(EvaluationFailure::Unsupported(
        super::super::UnsupportedKind::Function,
    ))?;
    if !function.valid_arity(arguments.len()) {
        return invalid_arity(evaluator, node, arguments);
    }
    match function {
        Function::Address => apply_address(evaluator, node, arguments),
        Function::Choose => apply_choose(evaluator, node, arguments),
        Function::HLookup | Function::Lookup | Function::Match | Function::VLookup => {
            // Validate complete ForceArray data descriptors before scalar key,
            // selector, or mode conversion can enter the search kernel.  A
            // list/3-D/vector shape refusal is a formula #VALUE! and consumes
            // no search cells on either the lifted or scalar path.
            if let Err(error) = search::preflight(evaluator, function, &arguments)? {
                // A known key error keeps its identity even when the data
                // descriptor is refused. Inspecting it must not project a
                // reference key or read cells on this refusal path.
                let error = match arguments.first() {
                    Some(RuntimeValue::Scalar(WorkingValue::Error(key_error))) => *key_error,
                    _ => error,
                };
                return Ok(formula_error(error));
            }
            if lifting::should_search(evaluator, function, &arguments)? {
                lifting::search(evaluator, node, name, function, &arguments)
            } else {
                search::apply(evaluator, node, name, arguments)
            }
        },
        Function::Index => apply_index(evaluator, node, arguments),
        Function::Indirect => apply_indirect(evaluator, node, arguments),
        Function::Offset => apply_offset(evaluator, node, arguments),
    }
}

fn formula_error<'expr>(error: ScalarError) -> RuntimeValue<'expr> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

fn invalid_arity<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let mut error = None;
    for argument in arguments {
        if let RuntimeValue::Scalar(WorkingValue::Error(value)) = argument {
            error.get_or_insert(value);
        }
    }
    if let Some(error) = error {
        return Ok(formula_error(error));
    }
    evaluator
        .scalar
        .charge_work(u64::try_from(node.child_count()).unwrap_or(u64::MAX))?;
    Ok(formula_error(ScalarError::Value))
}

fn apply_address<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if lifting::should_address(evaluator, &arguments)? {
        return lifting::address(evaluator, node, &arguments);
    }
    // The scalar path preserves ordinary implicit intersection.  Matrix
    // lifting above is entered only when one scalar parameter contributes a
    // non-singleton output shape.
    apply_address_scalar(evaluator, node, &mut arguments)
}

fn apply_address_scalar<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    _node: super::super::Node<'expr>,
    arguments: &mut Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let mut values = [None, None, None, None, None];
    for index in (0..arguments.len()).rev() {
        let argument = arguments.pop().ok_or(EvaluationFailure::InvalidExpression(
            "ADDRESS argument stack is incomplete",
        ))?;
        values[index] = Some(if matches!(argument, RuntimeValue::Missing) {
            argument
        } else {
            scalar_value(evaluator, argument)?
        });
    }
    if let Some(error) = first_formula_error(&values) {
        return Ok(formula_error(error));
    }

    let row = match integer_value(evaluator, values[0].take())? {
        Ok(Some(value)) => match positive_address_integer(value) {
            Ok(value) => value,
            Err(error) => return Ok(formula_error(error)),
        },
        Ok(None) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Ok(formula_error(error)),
    };
    let column = match integer_value(evaluator, values[1].take())? {
        Ok(Some(value)) => match positive_address_integer(value) {
            Ok(value) => value,
            Err(error) => return Ok(formula_error(error)),
        },
        Ok(None) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Ok(formula_error(error)),
    };
    let abs = match optional_number(evaluator, values[2].take(), 1.0)? {
        Ok(value) if value.is_finite() && value.fract() == 0.0 && (1.0..=4.0).contains(&value) => {
            value as u8
        },
        Ok(_) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Ok(formula_error(error)),
    };
    let a1 = match optional_logical(evaluator, values[3].take(), true)? {
        Ok(value) => value,
        Err(error) => return Ok(formula_error(error)),
    };
    let sheet = match values[4].take() {
        None | Some(RuntimeValue::Missing) | Some(RuntimeValue::Empty) => None,
        Some(RuntimeValue::Scalar(WorkingValue::Error(error))) => {
            return Ok(formula_error(error));
        },
        Some(RuntimeValue::Scalar(value)) => {
            Some(super::super::to_text(value, &mut evaluator.scalar)?)
        },
        Some(_) => return Ok(formula_error(ScalarError::Value)),
    };
    let result = lookup::format_address(
        &mut evaluator.scalar,
        row,
        column,
        abs,
        a1,
        sheet.as_ref().map(|value| value.text.as_ref()),
    )?;
    Ok(RuntimeValue::Scalar(WorkingValue::Text(result)))
}

fn apply_choose<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    _node: super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let index = arguments
        .drain(..1)
        .next()
        .ok_or(EvaluationFailure::InvalidExpression(
            "CHOOSE index is missing",
        ))?;
    let index = match integer_value(evaluator, Some(index))? {
        Ok(Some(index)) if index >= 1.0 => index as usize,
        Ok(Some(_)) | Ok(None) => return Ok(formula_error(ScalarError::Value)),
        Err(error) => return Ok(formula_error(error)),
    };
    let Some(branch) = arguments.into_iter().nth(index - 1) else {
        return Ok(formula_error(ScalarError::Value));
    };
    Ok(branch)
}

fn apply_index<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    _node: super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let data = arguments
        .drain(..1)
        .next()
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDEX data is missing",
        ))?;
    let row_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let column_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let area_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let row = index_selector(evaluator, row_value, 0, ScalarError::Value)?;
    let column = index_selector(evaluator, column_value, 0, ScalarError::Value)?;
    let area_number = index_selector(evaluator, area_value, 1, ScalarError::Reference)?;
    let (row, column, area_number) = match (row, column, area_number) {
        (Ok(row), Ok(column), Ok(area_number)) => (row, column, area_number),
        (Err(error), _, _) | (_, Err(error), _) | (_, _, Err(error)) => {
            return Ok(formula_error(error));
        },
    };
    if area_number == 0 {
        return Ok(formula_error(ScalarError::Reference));
    }

    match data {
        RuntimeValue::Areas(areas) => apply_index_areas(evaluator, areas, row, column, area_number),
        RuntimeValue::Array(array) if area_number == 1 => {
            apply_index_array(evaluator, array, row, column)
        },
        RuntimeValue::Array(_) => Ok(formula_error(ScalarError::Reference)),
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Ok(formula_error(error)),
        RuntimeValue::Scalar(_) => Ok(formula_error(ScalarError::Value)),
        RuntimeValue::Empty | RuntimeValue::Missing => Ok(formula_error(ScalarError::Value)),
        RuntimeValue::ScalarCell(area) => {
            let areas = scalar_cell_set(evaluator, area)?;
            apply_index_areas(evaluator, areas, row, column, area_number)
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn apply_index_areas<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    areas: RuntimeAreaSet<'expr>,
    row: usize,
    column: usize,
    area_number: usize,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if area_number == 0 || area_number > areas.records.len() {
        return Ok(formula_error(ScalarError::Reference));
    }
    let (start, end) = record_range(evaluator, &areas, area_number - 1)?;
    let selected = &areas.areas[start..end];
    let mut result = RuntimeAreaSet::derived(evaluator)?;
    // INDEX selects one logical record.  Even when the input is a reference
    // list, the selected descriptor is an ordinary reference thereafter.
    result.is_list = false;
    for area in selected.iter().copied() {
        evaluator.scalar.charge_work(1)?;
        let rect = slice_rect(area.rect, row, column)?;
        let Some(rect) = rect else {
            return Ok(formula_error(ScalarError::Reference));
        };
        result.push(
            RuntimeArea {
                sheet: area.sheet,
                sheet_index: area.sheet_index,
                rect,
            },
            evaluator,
        )?;
    }
    Ok(RuntimeValue::Areas(result.mark_scalar_result()))
}

fn apply_index_array<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    mut array: RuntimeArrayValue<'expr>,
    row: usize,
    column: usize,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some((row_start, row_end)) = indexed_extent(row, array.shape.rows()) else {
        return Ok(formula_error(ScalarError::Reference));
    };
    let Some((column_start, column_end)) = indexed_extent(column, array.shape.columns()) else {
        return Ok(formula_error(ScalarError::Reference));
    };
    if row == 0 && column == 0 {
        array.origin = None;
        array.preserve_scalar_result = false;
        return Ok(RuntimeValue::Array(array));
    }
    if row != 0 && column != 0 {
        evaluator.charge_cell_work(0)?;
        let source = row_start
            .checked_mul(array.shape.columns())
            .and_then(|offset| offset.checked_add(column_start))
            .ok_or(EvaluationFailure::InvalidExpression(
                "INDEX array offset overflows",
            ))?;
        let element = array
            .cells
            .get_mut(source)
            .ok_or(EvaluationFailure::InvalidExpression(
                "INDEX array element is missing",
            ))?;
        let selected = std::mem::replace(element, RuntimeElement::Empty);
        return evaluator.element_to_runtime(selected);
    }
    let rows = row_end - row_start;
    let columns = column_end - column_start;
    let shape = Shape::new(rows, columns)?;
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDEX array cell count overflows",
        ))?;
    let (mut output, reservation) = evaluator.new_element_vec(cells)?;
    for output_row in 0..rows {
        for output_column in 0..columns {
            let source = (row_start + output_row)
                .checked_mul(array.shape.columns())
                .and_then(|offset| offset.checked_add(column_start + output_column))
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDEX array offset overflows",
                ))?;
            evaluator.charge_cell_work(output.len())?;
            let element =
                array
                    .cells
                    .get_mut(source)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "INDEX array element is missing",
                    ))?;
            output.push(std::mem::replace(element, RuntimeElement::Empty));
        }
    }
    // A selected array has its own origin.  Retaining the input origin would
    // make a later scalar intersection use coordinates from the unsliced
    // source instead of the selected array geometry.
    evaluator.make_array(shape, output, reservation, None)
}

fn apply_offset<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    _node: super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let reference = arguments
        .drain(..1)
        .next()
        .ok_or(EvaluationFailure::InvalidExpression(
            "OFFSET reference is missing",
        ))?;
    let areas = match reference {
        RuntimeValue::Areas(areas) => areas,
        RuntimeValue::Scalar(WorkingValue::Error(error)) => return Ok(formula_error(error)),
        RuntimeValue::SourceReference => {
            return Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference,
            ));
        },
        _ => return Ok(formula_error(ScalarError::Value)),
    };
    if areas.is_list || areas.areas.is_empty() {
        return Ok(formula_error(ScalarError::Value));
    }
    let rows_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let columns_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let height_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let width_value = (!arguments.is_empty()).then(|| arguments.remove(0));
    let rows = signed_integer(evaluator, rows_value)?;
    let columns = signed_integer(evaluator, columns_value)?;
    let height = optional_positive_integer(evaluator, height_value, areas.areas[0].rect.rows())?;
    let width = optional_positive_integer(evaluator, width_value, areas.areas[0].rect.columns())?;
    let (rows, columns, height, width) = match (rows, columns, height, width) {
        (Ok(rows), Ok(columns), Ok(height), Ok(width)) => (rows, columns, height, width),
        (Err(error), _, _, _)
        | (_, Err(error), _, _)
        | (_, _, Err(error), _)
        | (_, _, _, Err(error)) => {
            return Ok(formula_error(error));
        },
    };
    let mut result = RuntimeAreaSet::derived(evaluator)?;
    for area in areas.areas.iter().copied() {
        let Some(row_start) = shift(area.rect.row_start, rows)? else {
            return Ok(formula_error(ScalarError::Reference));
        };
        let Some(column_start) = shift(area.rect.column_start, columns)? else {
            return Ok(formula_error(ScalarError::Reference));
        };
        let Some(row_end) = row_start.checked_add(height) else {
            return Ok(formula_error(ScalarError::Reference));
        };
        let Some(column_end) = column_start.checked_add(width) else {
            return Ok(formula_error(ScalarError::Reference));
        };
        let rect = match Rect::new(row_start, row_end, column_start, column_end) {
            Ok(rect) => rect,
            Err(_) => return Ok(formula_error(ScalarError::Reference)),
        };
        evaluator.scalar.charge_work(1)?;
        let extent = evaluator
            .resolver
            .sheet_extent(evaluator.sheet_name(area.sheet), evaluator.execution)?;
        evaluator.scalar.charge_work(0)?;
        let Some(extent) = extent else {
            return Ok(formula_error(ScalarError::Reference));
        };
        if rect.row_end > extent.rows() || rect.column_end > extent.columns() {
            return Ok(formula_error(ScalarError::Reference));
        }
        result.push(
            RuntimeArea {
                sheet: area.sheet,
                sheet_index: area.sheet_index,
                rect,
            },
            evaluator,
        )?;
    }
    Ok(RuntimeValue::Areas(result.mark_scalar_result()))
}

fn apply_indirect<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    _node: super::super::Node<'expr>,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let text = scalar_value(evaluator, arguments.remove(0))?;
    let text = match text {
        RuntimeValue::Empty => super::super::TextValue::borrowed(""),
        RuntimeValue::Scalar(value) => match lookup::text_argument(&mut evaluator.scalar, value)? {
            Ok(text) => text,
            Err(error) => return Ok(formula_error(error)),
        },
        _ => return Ok(formula_error(ScalarError::Value)),
    };
    let a1 = match optional_logical(evaluator, arguments.pop(), true)? {
        Ok(value) => value,
        Err(error) => return Ok(formula_error(error)),
    };
    let parsed = match lookup::parse_indirect(
        &mut evaluator.scalar,
        text.text.as_ref(),
        a1,
        evaluator.position.row,
        evaluator.position.column,
    )? {
        Ok(parsed) => parsed,
        Err(error) => return Ok(formula_error(error)),
    };
    let (reference, requires_origin, reservation) = parsed.into_parts();
    let result = match reference.as_ref() {
        Some(reference) if !requires_origin => indirect::adapt(evaluator, reference),
        Some(_) => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
        None => Ok(formula_error(ScalarError::Reference)),
    };
    // The descriptor borrows only canonical resolver names.  Release parser
    // storage after adaptation so the reservation never outlives its text.
    drop(reference);
    drop(reservation);
    result
}

fn scalar_cell_set<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    area: RuntimeArea<'expr>,
) -> EvaluationResult<RuntimeAreaSet<'expr>>
where
    R: Resolver + ?Sized,
{
    let mut result = RuntimeAreaSet::derived(evaluator)?;
    result.push(area, evaluator)?;
    Ok(result)
}

fn record_range<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    areas: &RuntimeAreaSet<'_>,
    record: usize,
) -> EvaluationResult<(usize, usize)> {
    let mut start = 0usize;
    for (index, value) in areas.records.iter().enumerate() {
        evaluator.scalar.charge_work(1)?;
        let planes = value.areas.iter().try_fold(0usize, |total, area| {
            evaluator.scalar.charge_work(1)?;
            total
                .checked_add(area.extent()[0])
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDEX record plane count overflows",
                ))
        })?;
        let end = start
            .checked_add(planes)
            .ok_or(EvaluationFailure::InvalidExpression(
                "INDEX record area range overflows",
            ))?;
        if index == record {
            if end > areas.areas.len() {
                return Err(EvaluationFailure::InvalidExpression(
                    "INDEX record areas are not contiguous",
                ));
            }
            return Ok((start, end));
        }
        start = end;
    }
    Err(EvaluationFailure::InvalidExpression(
        "INDEX area number is out of bounds",
    ))
}

fn slice_rect(rect: Rect, row: usize, column: usize) -> EvaluationResult<Option<Rect>> {
    let Some((row_start, row_end)) = indexed_extent(row, rect.rows()) else {
        return Ok(None);
    };
    let Some((column_start, column_end)) = indexed_extent(column, rect.columns()) else {
        return Ok(None);
    };
    Rect::new(
        rect.row_start + row_start,
        rect.row_start + row_end,
        rect.column_start + column_start,
        rect.column_start + column_end,
    )
    .map(Some)
}

fn indexed_extent(index: usize, length: usize) -> Option<(usize, usize)> {
    if index == 0 {
        Some((0, length))
    } else if index <= length {
        Some((index - 1, index))
    } else {
        None
    }
}

fn shift(origin: usize, offset: isize) -> EvaluationResult<Option<usize>> {
    if offset >= 0 {
        Ok(origin.checked_add(offset as usize))
    } else {
        Ok(origin.checked_sub(offset.unsigned_abs()))
    }
}

fn scalar_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    scalar_parameter(evaluator, value)
}

fn integer_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<Result<Option<f64>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(value) = value else {
        return Ok(Ok(None));
    };
    let value = scalar_parameter(evaluator, value)?;
    let value = match value {
        RuntimeValue::Empty | RuntimeValue::Missing => return Ok(Ok(Some(0.0))),
        RuntimeValue::Scalar(value) => value,
        _ => return Ok(Err(ScalarError::Value)),
    };
    match super::super::to_integer(value, &mut evaluator.scalar)? {
        Ok(value) => Ok(Ok(Some(value))),
        Err(error) => Ok(Err(error)),
    }
}

fn index_selector<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
    default: usize,
    negative_error: ScalarError,
) -> EvaluationResult<Result<usize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(value) = value else {
        return Ok(Ok(default));
    };
    if matches!(&value, RuntimeValue::Missing) {
        return Ok(Ok(default));
    }
    let value = integer_value(evaluator, Some(value))?;
    match value {
        Ok(Some(value)) if value < 0.0 => Ok(Err(negative_error)),
        // The exclusive power-of-two bound avoids rounded MAX admitting an
        // unrepresentable coordinate through a saturating float cast.
        Ok(Some(value)) if value < (2_f64).powi(usize::BITS as i32) => Ok(Ok(value as usize)),
        Ok(Some(_)) => Ok(Err(ScalarError::Reference)),
        Ok(None) => Ok(Err(ScalarError::Value)),
        Err(error) => Ok(Err(error)),
    }
}

fn signed_integer<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<Result<isize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    if matches!(value, None | Some(RuntimeValue::Missing)) {
        return Ok(Err(ScalarError::Value));
    }
    let bound = (2_f64).powi(isize::BITS as i32 - 1);
    match integer_value(evaluator, value)? {
        Ok(Some(value)) if value >= -bound && value < bound => Ok(Ok(value as isize)),
        Ok(Some(_)) => Ok(Err(ScalarError::Reference)),
        Ok(None) => Ok(Err(ScalarError::Value)),
        Err(error) => Ok(Err(error)),
    }
}

fn optional_positive_integer<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
    default: usize,
) -> EvaluationResult<Result<usize, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(value) = value else {
        return Ok(Ok(default));
    };
    if matches!(&value, RuntimeValue::Missing) {
        return Ok(Ok(default));
    }
    let value = index_selector(evaluator, Some(value), 0, ScalarError::Value)?;
    Ok(value.and_then(|value| {
        if value == 0 {
            Err(ScalarError::Value)
        } else {
            Ok(value)
        }
    }))
}

fn optional_number<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
    default: f64,
) -> EvaluationResult<Result<f64, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(value) = value else {
        return Ok(Ok(default));
    };
    if matches!(value, RuntimeValue::Missing) {
        return Ok(Ok(default));
    }
    let value = scalar_parameter(evaluator, value)?;
    match value {
        RuntimeValue::Empty => Ok(Ok(0.0)),
        RuntimeValue::Scalar(value) => super::super::to_integer(value, &mut evaluator.scalar),
        _ => Ok(Err(ScalarError::Value)),
    }
}

fn optional_logical<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: Option<RuntimeValue<'expr>>,
    default: bool,
) -> EvaluationResult<Result<bool, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(value) = value else {
        return Ok(Ok(default));
    };
    if matches!(value, RuntimeValue::Missing) {
        return Ok(Ok(default));
    }
    let value = scalar_parameter(evaluator, value)?;
    match value {
        RuntimeValue::Empty => Ok(Ok(false)),
        RuntimeValue::Scalar(value) => super::super::to_logical(value, &mut evaluator.scalar),
        _ => Ok(Err(ScalarError::Value)),
    }
}

fn scalar_parameter<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    evaluator.matrix_scalar_parameter(value)
}

fn first_formula_error(values: &[Option<RuntimeValue<'_>>; 5]) -> Option<ScalarError> {
    values.iter().find_map(|value| match value.as_ref()? {
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Some(*error),
        _ => None,
    })
}

/// Convert ADDRESS row/column arguments with the checked positive integer
/// profile used at the value boundary. In particular, reject values at or
/// above the first unrepresentable `usize` instead of relying on Rust's
/// saturating float-to-integer cast.
fn positive_address_integer(value: f64) -> Result<usize, ScalarError> {
    if !value.is_finite() || value < 1.0 || value.fract() != 0.0 {
        return Err(ScalarError::Value);
    }
    if value >= (usize::MAX as u128 + 1) as f64 {
        return Err(ScalarError::Value);
    }
    Ok(value as usize)
}
