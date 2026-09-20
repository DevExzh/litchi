//! Scalar-kernel bridge for the value VM. No AST is re-evaluated here.

use super::super::{
    EvaluationFailure, EvaluationResult, Evaluator, InfixOperator, Node, PostfixOperator,
    PrefixOperator, ScalarError, TextValue, UnsupportedKind, WorkingValue,
};

/// Keep Empty distinct until the operand's required type is known.
pub(super) enum Slot<'a> {
    Empty,
    Value(WorkingValue<'a>),
}

#[derive(Clone, Copy)]
pub(super) enum EmptyTarget {
    Number,
    Logical,
    Text,
}

impl<'a> Slot<'a> {
    pub(super) fn coerce_empty(self, target: EmptyTarget) -> WorkingValue<'a> {
        match self {
            Self::Value(value) => value,
            Self::Empty => match target {
                EmptyTarget::Number => WorkingValue::Number(0.0),
                EmptyTarget::Logical => WorkingValue::Logical(false),
                EmptyTarget::Text => WorkingValue::Text(TextValue::borrowed("")),
            },
        }
    }
}

pub(super) fn prefix<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    operator: PrefixOperator,
    operand: Slot<'a>,
) -> EvaluationResult<Slot<'a>> {
    // Unary + accepts Any and preserves its operand, including Empty.
    if operator == PrefixOperator::Plus {
        return Ok(operand);
    }
    super::super::apply_prefix(
        operator,
        operand.coerce_empty(EmptyTarget::Number),
        evaluator,
    )
    .map(Slot::Value)
}

pub(super) fn postfix<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    operator: PostfixOperator,
    operand: Slot<'a>,
) -> EvaluationResult<Slot<'a>> {
    super::super::apply_postfix(
        operator,
        operand.coerce_empty(EmptyTarget::Number),
        evaluator,
    )
    .map(Slot::Value)
}

pub(super) fn infix<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    operator: InfixOperator,
    left: Slot<'a>,
    right: Slot<'a>,
) -> EvaluationResult<Slot<'a>> {
    // Preserve formula-error precedence before applying the Empty profile.
    if let Slot::Value(WorkingValue::Error(error)) = &left {
        return Ok(Slot::Value(WorkingValue::Error(*error)));
    }
    if let Slot::Value(WorkingValue::Error(error)) = &right {
        return Ok(Slot::Value(WorkingValue::Error(*error)));
    }
    let left_empty = matches!(left, Slot::Empty);
    let right_empty = matches!(right, Slot::Empty);
    if left_empty || right_empty {
        let result = match operator {
            InfixOperator::Equal => Some(WorkingValue::Logical(left_empty && right_empty)),
            InfixOperator::NotEqual => Some(WorkingValue::Logical(left_empty != right_empty)),
            InfixOperator::Less
            | InfixOperator::LessEqual
            | InfixOperator::Greater
            | InfixOperator::GreaterEqual => Some(WorkingValue::Error(ScalarError::Value)),
            _ => None,
        };
        if let Some(value) = result {
            return Ok(Slot::Value(value));
        }
    }
    let target = if operator == InfixOperator::Concatenate {
        EmptyTarget::Text
    } else {
        EmptyTarget::Number
    };
    evaluator
        .apply_infix(
            operator,
            left.coerce_empty(target),
            right.coerce_empty(target),
        )
        .map(Slot::Value)
}

/// Parameter conversion for the eager scalar function catalog.
/// TextOrNumber radix arguments use the explicit empty-Text profile.
pub(super) fn argument<'a>(name: &str, index: usize, slot: Slot<'a>) -> WorkingValue<'a> {
    let text_first = index == 0
        && [
            "ARABIC", "DECIMAL", "BIN2DEC", "BIN2HEX", "BIN2OCT", "HEX2BIN", "HEX2DEC", "HEX2OCT",
            "OCT2BIN", "OCT2DEC", "OCT2HEX",
        ]
        .iter()
        .any(|function| name.eq_ignore_ascii_case(function));
    let text_empty =
        matches!(slot, Slot::Empty) && super::super::text::argument_is_text(name, index);
    let target = if text_first || text_empty {
        EmptyTarget::Text
    } else if name.eq_ignore_ascii_case("NOT") || name.eq_ignore_ascii_case("XOR") {
        EmptyTarget::Logical
    } else {
        EmptyTarget::Number
    };
    slot.coerce_empty(target)
}

/// Apply TEXT's Scalar type gate after every argument has been read. A source
/// formula error takes precedence over the generated Empty type refusal.
pub(super) fn refuse_empty_text_value(arguments: &mut [WorkingValue<'_>], first_empty: bool) {
    if first_empty
        && !arguments
            .iter()
            .any(|value| matches!(value, WorkingValue::Error(_)))
        && let Some(first) = arguments.first_mut()
    {
        *first = WorkingValue::Error(ScalarError::Value);
    }
}

/// Execute an eager scalar function using the scalar VM's existing stack.
///
/// The caller resolves array/reference projection and Empty conversion first,
/// and retains any argument-storage reservations through this call. Arguments
/// may be produced fallibly one at a time; this bridge adds no staging vector.
/// Lazy functions and sequence aggregation belong to the value VM.
pub(super) fn eager<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
    arguments: impl IntoIterator<Item = EvaluationResult<WorkingValue<'a>>>,
) -> EvaluationResult<WorkingValue<'a>> {
    evaluator.charge_work(1)?;
    if !evaluator.values.is_empty() || !evaluator.frames.is_empty() {
        return Err(EvaluationFailure::InvalidExpression(
            "scalar bridge requires empty kernel value and frame stacks",
        ));
    }
    if !matches!(node.kind(), super::super::Kind::Function { name: actual } if actual.eq_ignore_ascii_case(name))
    {
        return Err(EvaluationFailure::InvalidExpression(
            "scalar bridge function does not match the expression node",
        ));
    }
    if ["IF", "IFERROR", "IFNA", "AND", "OR"]
        .iter()
        .any(|function| name.eq_ignore_ascii_case(function))
    {
        return Err(EvaluationFailure::InvalidExpression(
            "lazy or sequence function reached scalar bridge",
        ));
    }
    if ![
        "TRUE",
        "FALSE",
        "XOR",
        "NOT",
        "BITAND",
        "BITOR",
        "BITXOR",
        "BITLSHIFT",
        "BITRSHIFT",
    ]
    .iter()
    .any(|function| name.eq_ignore_ascii_case(function))
        && !super::super::radix::is_radix_function(name)
        && !super::super::roman::is_roman_function(name)
        && !super::super::complex::is_complex_function(name)
        && !super::super::rounding::is_rounding_function(name)
        && !super::super::trigonometry::is_trigonometric_function(name)
        && !super::super::elementary::is_elementary_function(name)
        && !super::super::date_time::is_date_time_function(name)
        && ![
            "COMBIN",
            "COMBINA",
            "FACT",
            "FACTDOUBLE",
            "EVEN",
            "ODD",
            "DELTA",
            "GESTEP",
        ]
        .iter()
        .any(|function| name.eq_ignore_ascii_case(function))
        && !super::super::is_conditional_aggregate_function(name)
        && !super::super::text::is_text_function(name)
    {
        return Err(EvaluationFailure::Unsupported(UnsupportedKind::Function));
    }

    let result = (|| {
        for argument in arguments {
            evaluator.charge_work(1)?;
            if evaluator.values.len() >= node.child_count() {
                return Err(EvaluationFailure::InvalidExpression(
                    "scalar bridge received too many arguments",
                ));
            }
            evaluator.push_value(argument?)?;
        }
        if evaluator.values.len() != node.child_count() {
            return Err(EvaluationFailure::InvalidExpression(
                "scalar bridge received too few arguments",
            ));
        }
        if node.child_count() == 0
            && (name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE"))
        {
            return Ok(WorkingValue::Logical(name.eq_ignore_ascii_case("TRUE")));
        }
        evaluator.apply_function(node, name)?;
        if evaluator.values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "scalar kernel did not produce exactly one value",
            ));
        }
        evaluator.pop_value()
    })();
    // Partial argument failures must release retained text and leave the
    // reusable kernel stack ready for the next operation.
    evaluator.values.clear();
    result
}

pub(super) fn logical(
    evaluator: &mut Evaluator<'_, '_, '_>,
    operand: Slot<'_>,
) -> EvaluationResult<Result<bool, ScalarError>> {
    super::super::to_logical(operand.coerce_empty(EmptyTarget::Logical), evaluator)
}

#[cfg(test)]
mod tests {
    use super::super::super::{EvaluationContext, EvaluationLimits, Expression};
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Profile};
    use std::{
        num::{NonZeroU64, NonZeroUsize},
        sync::Arc,
    };

    fn execution() -> ExecutionContext {
        execution_with_cancellation().1
    }

    fn execution_with_cancellation() -> (CancellationSource, ExecutionContext) {
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionContext::new(
            Budget::root(
                Arc::from("value-scalar-bridge-test"),
                litchi_core::Limits::for_profile(Profile::Desktop),
            ),
            token,
            ExecutionLimits::new(
                NonZeroUsize::MIN,
                NonZeroUsize::MIN,
                NonZeroU64::new(1 << 30).unwrap(),
                1 << 20,
            )
            .unwrap(),
        );
        (source, execution)
    }

    #[test]
    fn failed_argument_releases_kernel_values_and_allows_reuse() {
        let execution = execution();
        let context = EvaluationContext::new(&execution);
        let expression = Expression::parse("=BITAND(1;2)").unwrap();
        let mut evaluator = Evaluator::new(
            &expression,
            &context,
            EvaluationLimits::default(),
            execution.budget().clone(),
        );
        evaluator.push_value(WorkingValue::Number(0.0)).unwrap();
        evaluator.pop_value().unwrap();
        let retained_stack_bytes = execution.budget().used(litchi_core::Resource::Memory);
        let text = TextValue::owned(
            "1".to_owned(),
            evaluator.reserve_storage(1, "test argument text").unwrap(),
        );
        let error = eager(
            &mut evaluator,
            expression.root(),
            "BITAND",
            [
                Ok(WorkingValue::Text(text)),
                Err(EvaluationFailure::Cancelled),
            ],
        )
        .err()
        .unwrap();
        assert!(matches!(error, EvaluationFailure::Cancelled));
        assert!(evaluator.values.is_empty());
        assert_eq!(
            execution.budget().used(litchi_core::Resource::Memory),
            retained_stack_bytes
        );
        let value = eager(
            &mut evaluator,
            expression.root(),
            "BITAND",
            [Ok(WorkingValue::Number(7.0)), Ok(WorkingValue::Number(3.0))],
        )
        .unwrap();
        assert!(matches!(value, WorkingValue::Number(3.0)));
        assert!(evaluator.values.is_empty());
    }

    #[test]
    fn empty_keeps_any_identity_and_uses_required_conversion() {
        let execution = execution();
        let context = EvaluationContext::new(&execution);
        let expression = Expression::parse("=1").unwrap();
        let mut evaluator = Evaluator::new(
            &expression,
            &context,
            EvaluationLimits::default(),
            execution.budget().clone(),
        );
        assert!(matches!(
            prefix(&mut evaluator, PrefixOperator::Plus, Slot::Empty).unwrap(),
            Slot::Empty
        ));
        assert_eq!(logical(&mut evaluator, Slot::Empty).unwrap(), Ok(false));
        assert!(
            matches!(Slot::Empty.coerce_empty(EmptyTarget::Text), WorkingValue::Text(text) if text.text.is_empty())
        );
        assert!(
            matches!(prefix(&mut evaluator, PrefixOperator::Minus, Slot::Empty).unwrap(), Slot::Value(WorkingValue::Number(value)) if value == 0.0)
        );
    }

    #[test]
    fn eager_bridge_handles_constants_but_rejects_lazy_dispatch() {
        let execution = execution();
        let context = EvaluationContext::new(&execution);
        let expression = Expression::parse("=TRUE()").unwrap();
        let mut evaluator = Evaluator::new(
            &expression,
            &context,
            EvaluationLimits::default(),
            execution.budget().clone(),
        );
        assert!(matches!(
            eager(&mut evaluator, expression.root(), "TRUE", []).unwrap(),
            WorkingValue::Logical(true)
        ));
        assert!(matches!(
            eager(&mut evaluator, expression.root(), "IF", []),
            Err(EvaluationFailure::InvalidExpression(_))
        ));
        assert!(evaluator.frames.is_empty());
    }

    #[test]
    fn empty_comparison_and_text_parameter_profiles_are_explicit() {
        let execution = execution();
        let context = EvaluationContext::new(&execution);
        let expression = Expression::parse("=1").unwrap();
        let mut evaluator = Evaluator::new(
            &expression,
            &context,
            EvaluationLimits::default(),
            execution.budget().clone(),
        );
        assert!(matches!(
            infix(
                &mut evaluator,
                InfixOperator::Equal,
                Slot::Empty,
                Slot::Empty
            )
            .unwrap(),
            Slot::Value(WorkingValue::Logical(true))
        ));
        for other in [
            WorkingValue::Number(0.0),
            WorkingValue::Logical(false),
            WorkingValue::Text(TextValue::borrowed("")),
        ] {
            assert!(matches!(
                infix(
                    &mut evaluator,
                    InfixOperator::Equal,
                    Slot::Empty,
                    Slot::Value(other)
                )
                .unwrap(),
                Slot::Value(WorkingValue::Logical(false))
            ));
        }
        assert!(matches!(
            infix(
                &mut evaluator,
                InfixOperator::Less,
                Slot::Empty,
                Slot::Empty
            )
            .unwrap(),
            Slot::Value(WorkingValue::Error(ScalarError::Value))
        ));
        assert!(matches!(
            infix(
                &mut evaluator,
                InfixOperator::Equal,
                Slot::Empty,
                Slot::Value(WorkingValue::Error(ScalarError::NotAvailable))
            )
            .unwrap(),
            Slot::Value(WorkingValue::Error(ScalarError::NotAvailable))
        ));
        assert!(
            matches!(infix(&mut evaluator, InfixOperator::Concatenate, Slot::Empty, Slot::Value(WorkingValue::Text(TextValue::borrowed("tail")))).unwrap(), Slot::Value(WorkingValue::Text(text)) if text.text == "tail")
        );
        assert!(
            matches!(postfix(&mut evaluator, PostfixOperator::Percent, Slot::Empty).unwrap(), Slot::Value(WorkingValue::Number(value)) if value == 0.0)
        );
        assert!(
            matches!(argument("DECIMAL", 0, Slot::Empty), WorkingValue::Text(text) if text.text.is_empty())
        );
        assert!(matches!(
            argument("DECIMAL", 1, Slot::Empty),
            WorkingValue::Number(0.0)
        ));
        assert!(
            matches!(argument("BIN2DEC", 0, Slot::Empty), WorkingValue::Text(text) if text.text.is_empty())
        );
    }

    #[test]
    fn bridge_rejects_pending_frames_and_checks_constant_cancellation() {
        let (source, execution) = execution_with_cancellation();
        let context = EvaluationContext::new(&execution);
        let expression = Expression::parse("=TRUE()").unwrap();
        let mut evaluator = Evaluator::new(
            &expression,
            &context,
            EvaluationLimits::default(),
            execution.budget().clone(),
        );
        evaluator
            .push_frame(super::super::super::Frame::Visit(expression.root()))
            .unwrap();
        assert!(matches!(
            eager(&mut evaluator, expression.root(), "TRUE", []),
            Err(EvaluationFailure::InvalidExpression(_))
        ));
        assert_eq!(evaluator.frames.len(), 1);
        evaluator.frames.clear();
        assert!(matches!(
            eager(&mut evaluator, expression.root(), "FALSE", []),
            Err(EvaluationFailure::InvalidExpression(_))
        ));
        source.cancel();
        assert!(matches!(
            eager(&mut evaluator, expression.root(), "TRUE", []),
            Err(EvaluationFailure::Cancelled)
        ));
        assert!(evaluator.frames.is_empty());
        assert!(evaluator.values.is_empty());
    }

    #[test]
    fn lazy_and_sequence_functions_do_not_consume_bridge_arguments() {
        let execution = execution();
        let context = EvaluationContext::new(&execution);
        for source in [
            "=IF(1;2;3)",
            "=IFERROR(1;2)",
            "=IFNA(1;2)",
            "=AND(1;2)",
            "=OR(1;2)",
        ] {
            let expression = Expression::parse(source).unwrap();
            let super::super::super::Kind::Function { name } = expression.root().kind() else {
                panic!("function fixture")
            };
            let mut evaluator = Evaluator::new(
                &expression,
                &context,
                EvaluationLimits::default(),
                execution.budget().clone(),
            );
            let calls = std::cell::Cell::new(0usize);
            let arguments = std::iter::from_fn(|| {
                calls.set(calls.get() + 1);
                Some(Ok(WorkingValue::Number(1.0)))
            });
            assert!(matches!(
                eager(&mut evaluator, expression.root(), name, arguments),
                Err(EvaluationFailure::InvalidExpression(_))
            ));
            assert_eq!(calls.get(), 0);
            assert!(evaluator.values.is_empty());
            assert!(evaluator.frames.is_empty());
        }
    }
}
