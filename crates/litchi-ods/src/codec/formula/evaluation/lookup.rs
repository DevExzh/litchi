//! Shared scalar kernels and catalog for the OpenFormula lookup family.
//!
//! Resolver-backed traversal belongs to `evaluation/value/lookup.rs`.  This
//! module owns only function identity, argument scheduling metadata, bounded
//! scalar conversions, ADDRESS text construction, INDIRECT lexical adapters,
//! and the comparison rules shared by exact and approximate searches.

mod address;
mod indirect;
mod search;

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextValue, UnsupportedKind,
    WorkingValue,
};

pub(super) use address::format as format_address;
pub(super) use indirect::{parse_indirect, parse_indirect_without_origin};
pub(super) use search::{SearchValue, approximate_type_compatible, compare_values, exact_match};

/// The nine functions in OpenFormula §6.14 implemented by the lookup batch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Function {
    Address,
    Choose,
    HLookup,
    Index,
    Indirect,
    Lookup,
    Match,
    Offset,
    VLookup,
}

impl Function {
    /// Resolve one case-insensitive function name without allocating.
    pub(super) fn from_name(name: &str) -> Option<Self> {
        let first = name.as_bytes().first()?.to_ascii_uppercase();
        match (first, name.len()) {
            (b'A', 7) if name.eq_ignore_ascii_case("ADDRESS") => Some(Self::Address),
            (b'C', 6) if name.eq_ignore_ascii_case("CHOOSE") => Some(Self::Choose),
            (b'H', 7) if name.eq_ignore_ascii_case("HLOOKUP") => Some(Self::HLookup),
            (b'I', 5) if name.eq_ignore_ascii_case("INDEX") => Some(Self::Index),
            (b'I', 8) if name.eq_ignore_ascii_case("INDIRECT") => Some(Self::Indirect),
            (b'L', 6) if name.eq_ignore_ascii_case("LOOKUP") => Some(Self::Lookup),
            (b'M', 5) if name.eq_ignore_ascii_case("MATCH") => Some(Self::Match),
            (b'O', 6) if name.eq_ignore_ascii_case("OFFSET") => Some(Self::Offset),
            (b'V', 7) if name.eq_ignore_ascii_case("VLOOKUP") => Some(Self::VLookup),
            _ => None,
        }
    }

    /// Check the exact OpenFormula arity envelope. Optional omitted slots are
    /// represented by fewer AST children; explicit missing slots stay within
    /// the same envelope and are interpreted by the caller.
    pub(super) const fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Address => count >= 2 && count <= 5,
            Self::Choose => count >= 2,
            Self::HLookup | Self::VLookup => count >= 3 && count <= 4,
            Self::Index => count >= 1 && count <= 4,
            Self::Indirect => count >= 1 && count <= 2,
            Self::Lookup => count >= 2 && count <= 3,
            Self::Match => count >= 2 && count <= 3,
            Self::Offset => count >= 3 && count <= 5,
        }
    }

    /// Whether an argument must retain a complete ForceArray/reference
    /// descriptor before scalar publication.
    pub(super) const fn matrix_argument(self, index: usize) -> bool {
        match self {
            Self::HLookup | Self::VLookup => index == 1,
            Self::Index => index == 0,
            Self::Lookup => index == 1 || index == 2,
            Self::Match => index == 1,
            Self::Offset => index == 0,
            Self::Address | Self::Choose | Self::Indirect => false,
        }
    }

    /// Whether the function selects and evaluates only one of its children.
    pub(super) const fn is_lazy(self) -> bool {
        matches!(self, Self::Choose)
    }

    /// Whether an argument is the input descriptor/reference for a function.
    /// The table lookup family accepts `Reference|Array`; OFFSET accepts a
    /// strict Reference; CHOOSE and ADDRESS have no descriptor argument.
    pub(super) const fn reference_argument(self, index: usize) -> bool {
        match self {
            Self::HLookup | Self::VLookup => index == 1,
            Self::Index => index == 0,
            Self::Lookup => index == 1 || index == 2,
            Self::Match => index == 1,
            Self::Offset => index == 0,
            Self::Address | Self::Choose | Self::Indirect => false,
        }
    }
}

/// Return whether `name` belongs to the lookup family.
pub(super) fn is_lookup_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

/// A scalar conversion result with formula errors retained as values.
pub(super) fn integer_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match super::to_integer(value, evaluator)? {
        Ok(value) => Ok(Ok(value)),
        Err(error) => Ok(Err(error)),
    }
}

/// Convert a scalar value to a logical flag using the shared scalar profile.
pub(super) fn logical_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<bool, ScalarError>> {
    super::to_logical(value, evaluator)
}

/// Convert a scalar value to borrowed or bounded-owned Text using the shared
/// evaluator policy.
pub(super) fn text_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    if let WorkingValue::Error(error) = value {
        return Ok(Err(error));
    }
    match super::to_text(value, evaluator) {
        Ok(value) => Ok(Ok(value)),
        Err(error) => Err(error),
    }
}

/// Apply context-free lookup cases through the scalar evaluator's eager
/// stack. Resolver-dependent functions are rejected as a typed capability;
/// the value VM owns their descriptor and ForceArray behavior.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = Function::from_name(name)
        .ok_or(EvaluationFailure::Unsupported(UnsupportedKind::Function))?;
    if !function.valid_arity(node.child_count()) {
        return evaluator.finish_invalid_arity(node);
    }
    match function {
        Function::Address => apply_address(evaluator, node),
        // Scalar evaluation has a dedicated lazy CHOOSE continuation in the
        // parent evaluator. The value VM owns its branch selector; reaching
        // this eager dispatcher is an evaluator invariant failure.
        Function::Choose => Err(EvaluationFailure::InvalidExpression(
            "CHOOSE reached the eager lookup dispatcher",
        )),
        Function::Indirect => apply_indirect(evaluator, node),
        Function::HLookup
        | Function::Index
        | Function::Lookup
        | Function::Match
        | Function::Offset
        | Function::VLookup => apply_context_dependent(evaluator, node),
    }
}

fn apply_context_dependent<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    // A scalar literal cannot satisfy a Reference|Array/Reference parameter.
    // Preserve an already-produced formula error as a value before reporting
    // the ordinary shape refusal; typed provider/resource errors never reach
    // this scalar path because the scheduler has no resolver.
    let mut first_error = None;
    for _ in 0..node.child_count() {
        if let WorkingValue::Error(error) = evaluator.pop_value()? {
            // Arguments are popped from right to left; replacing here leaves
            // the leftmost formula error after the complete suffix is seen.
            first_error = Some(error);
        }
    }
    evaluator.push_value(WorkingValue::Error(
        first_error.unwrap_or(ScalarError::Value),
    ))
}

fn apply_indirect<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    let count = node.child_count();
    let mut values: [Option<WorkingValue<'a>>; 2] = std::array::from_fn(|_| None);
    for index in (0..count).rev() {
        values[index] = Some(evaluator.pop_value()?);
    }
    let mut first_error = None;
    for (index, value) in values.iter().enumerate().take(count) {
        if node.child(index).is_some_and(|child| child.is_missing()) {
            continue;
        }
        if let Some(WorkingValue::Error(error)) = value {
            first_error.get_or_insert(*error);
        }
    }
    if let Some(error) = first_error {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let a1 = match node.child(1) {
        None => true,
        Some(child) if child.is_missing() => true,
        Some(_) => {
            let value = values[1]
                .take()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDIRECT A1 argument is missing from the scalar stack",
                ))?;
            match logical_argument(evaluator, value)? {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            }
        },
    };
    let value = values[0]
        .take()
        .ok_or(EvaluationFailure::InvalidExpression(
            "INDIRECT text argument is missing from the scalar stack",
        ))?;
    let text = match text_argument(evaluator, value)? {
        Ok(text) => text,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    // The scalar evaluator has no Position/Resolver, but parsing still
    // validates malformed text before returning the capability refusal.
    match parse_indirect_without_origin(evaluator, text.text.as_ref(), a1)? {
        Ok(parsed) => {
            drop(parsed);
            Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
        },
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn apply_address<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<()> {
    let count = node.child_count();
    let mut values: [Option<WorkingValue<'a>>; 5] = std::array::from_fn(|_| None);
    for index in (0..count).rev() {
        values[index] = Some(evaluator.pop_value()?);
    }
    let mut first_error = None;
    for (index, value) in values.iter().enumerate().take(count) {
        if node.child(index).is_some_and(|child| child.is_missing()) {
            continue;
        }
        if let Some(WorkingValue::Error(error)) = value {
            first_error.get_or_insert(*error);
        }
    }
    if let Some(error) = first_error {
        return evaluator.push_value(WorkingValue::Error(error));
    }

    let row = match values[0].take() {
        Some(value) => match integer_argument(evaluator, value)? {
            Ok(value) => match positive_integer(value) {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            },
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => {
            return Err(EvaluationFailure::InvalidExpression(
                "ADDRESS row argument is missing",
            ));
        },
    };
    let column = match values[1].take() {
        Some(value) => match integer_argument(evaluator, value)? {
            Ok(value) => match positive_integer(value) {
                Ok(value) => value,
                Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
            },
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
        None => {
            return Err(EvaluationFailure::InvalidExpression(
                "ADDRESS column argument is missing",
            ));
        },
    };
    let abs_value = if node.child(2).is_some_and(|child| child.is_missing()) {
        None
    } else {
        values[2].take()
    };
    let abs = match optional_number(evaluator, abs_value, 1.0)? {
        Ok(value) => match positive_integer(value) {
            Ok(value) if value <= 4 => value as u8,
            _ => return evaluator.push_value(WorkingValue::Error(ScalarError::Value)),
        },
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let a1_value = if node.child(3).is_some_and(|child| child.is_missing()) {
        None
    } else {
        values[3].take()
    };
    let a1 = match optional_logical(evaluator, a1_value, true)? {
        Ok(value) => value,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    let sheet_value = if node.child(4).is_some_and(|child| child.is_missing()) {
        None
    } else {
        values[4].take()
    };
    let sheet = match sheet_value {
        None => None,
        Some(WorkingValue::Error(error)) => {
            return evaluator.push_value(WorkingValue::Error(error));
        },
        Some(value) => match text_argument(evaluator, value)? {
            Ok(text) => Some(text),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
    };
    let result = format_address(
        evaluator,
        row,
        column,
        abs,
        a1,
        sheet.as_ref().map(|text| text.text.as_ref()),
    )?;
    evaluator.push_value(WorkingValue::Text(result))
}

fn optional_number<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: Option<WorkingValue<'a>>,
    default: f64,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match value {
        None => Ok(Ok(default)),
        Some(value) => integer_argument(evaluator, value),
    }
}

fn optional_logical<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: Option<WorkingValue<'a>>,
    default: bool,
) -> EvaluationResult<Result<bool, ScalarError>> {
    match value {
        None => Ok(Ok(default)),
        Some(value) => logical_argument(evaluator, value),
    }
}

fn positive_integer(value: f64) -> Result<usize, ScalarError> {
    if !value.is_finite() || value < 1.0 || value.fract() != 0.0 {
        return Err(ScalarError::Value);
    }
    if value > usize::MAX as f64 {
        return Err(ScalarError::Number);
    }
    Ok(value as usize)
}
