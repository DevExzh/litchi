//! Scalar OpenFormula text functions.
//!
//! The scalar evaluator keeps text values as borrowed UTF-8 whenever an
//! operation can return a source slice.  Functions which have to construct a
//! result use the same fallible storage and work accounting as the rest of
//! the evaluator.  Matrix/reference text evaluation remains in the value VM;
//! this module only consumes the eagerly evaluated scalar argument tail.

mod bytes;
mod core;
mod format;
mod fraction;
mod search;
mod unicode;

/// Iterate the pinned Unicode C+F case-fold mapping for a lookup comparator.
/// The lookup family shares this table but keeps its own fixed-state iterator;
/// exposing the mapping here avoids a second Unicode data copy in that module.
pub(super) fn case_fold(value: char) -> impl Iterator<Item = char> {
    unicode::case_fold(value).iter()
}
mod width;

use super::{EvaluationFailure, EvaluationResult, Evaluator, Node, TextValue, WorkingValue};

/// Return whether `name` belongs to the sections 6.7 or 6.20 text families.
pub(super) fn is_text_function(name: &str) -> bool {
    bytes::is_byte_function(name)
        || core::is_core_function(name)
        || matches!(name.len(), 3..=10)
            && [
                "ASC", "CHAR", "CODE", "DOLLAR", "FIXED", "JIS", "TEXT", "UNICHAR", "UNICODE",
            ]
            .iter()
            .any(|function| name.eq_ignore_ascii_case(function))
}

/// Identify arguments whose Empty value is the Text target in the value VM.
///
/// The scalar evaluator has no Empty variant: an omitted AST slot is handled
/// by each function's optional-argument rule.  The value VM uses this
/// classifier when it has to coerce a matrix Empty to a scalar function's
/// declared Text parameter.
pub(super) fn argument_is_text(name: &str, index: usize) -> bool {
    if bytes::argument_is_text(name, index) || core::argument_is_text(name, index) {
        return true;
    }
    (name.eq_ignore_ascii_case("ASC")
        || name.eq_ignore_ascii_case("JIS")
        || name.eq_ignore_ascii_case("CODE")
        || name.eq_ignore_ascii_case("UNICODE"))
        && index == 0
        || name.eq_ignore_ascii_case("TEXT") && index == 1
}

/// Apply a scalar text function after eager argument evaluation.
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    if bytes::is_byte_function(name) {
        return bytes::apply(evaluator, node, name);
    }
    if core::is_core_function(name) {
        return core::apply(evaluator, node, name);
    }
    if is_width_function(name) {
        return apply_width(evaluator, node, name);
    }
    if format::is_format_function(name) {
        return format::apply(evaluator, node, name);
    }
    Err(EvaluationFailure::Unsupported(
        super::UnsupportedKind::Function,
    ))
}

fn is_width_function(name: &str) -> bool {
    name.eq_ignore_ascii_case("ASC") || name.eq_ignore_ascii_case("JIS")
}

fn apply_width<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = core::first_formula_error(evaluator, node, 1)? {
        core::discard_tail(evaluator, 1)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    core::reverse_value_tail(evaluator, 1)?;
    if core::has_complex_argument(evaluator, 1) {
        core::discard_tail(evaluator, 1)?;
        return evaluator.push_value(WorkingValue::Error(super::ScalarError::Value));
    }
    let value = evaluator.pop_value()?;
    let text = match value {
        WorkingValue::Error(error) => WorkingValue::Error(error),
        value => match super::to_text(value, evaluator) {
            Ok(text) => return apply_width_text(evaluator, text, name),
            Err(error) => return Err(error),
        },
    };
    evaluator.push_value(text)
}

fn apply_width_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: TextValue<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let source = text.text.as_ref();
    core::charge_text(evaluator, source)?;
    if source.is_empty() {
        return evaluator.push_value(WorkingValue::Text(text));
    }

    let mut output_len = 0usize;
    let mut changed = false;
    let mut iterator = source.char_indices().peekable();
    let mut iteration = 0usize;
    while let Some((_, character)) = iterator.next() {
        core::checkpoint(evaluator, iteration)?;
        iteration = iteration.saturating_add(1);
        let next = iterator.peek().map(|(_, character)| *character);
        let (mapped, suffix, consumed) = if name.eq_ignore_ascii_case("ASC") {
            let (mapped, suffix) = width::asc(character);
            (mapped, suffix, false)
        } else {
            let (mapped, consumed) = width::jis(character, next);
            (mapped, None, consumed)
        };
        if mapped != character || suffix.is_some() || consumed {
            changed = true;
        }
        output_len = output_len
            .checked_add(mapped.len_utf8())
            .ok_or_else(|| core::text_limit(evaluator, usize::MAX))?;
        if let Some(suffix) = suffix {
            output_len = output_len
                .checked_add(suffix.len_utf8())
                .ok_or_else(|| core::text_limit(evaluator, usize::MAX))?;
        }
        if consumed {
            iterator.next();
        }
    }

    if !changed && output_len == source.len() {
        return evaluator.push_value(WorkingValue::Text(text));
    }
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = core::new_output(evaluator, output_len)?;
    let mut iterator = source.char_indices().peekable();
    let mut iteration = 0usize;
    while let Some((_, character)) = iterator.next() {
        core::checkpoint(evaluator, iteration)?;
        iteration = iteration.saturating_add(1);
        let next = iterator.peek().map(|(_, character)| *character);
        let (mapped, suffix, consumed) = if name.eq_ignore_ascii_case("ASC") {
            let (mapped, suffix) = width::asc(character);
            (mapped, suffix, false)
        } else {
            let (mapped, consumed) = width::jis(character, next);
            (mapped, None, consumed)
        };
        output.push(mapped);
        if let Some(suffix) = suffix {
            output.push(suffix);
        }
        if consumed {
            iterator.next();
        }
        evaluator.charge_work(1)?;
    }
    evaluator.push_value(WorkingValue::Text(TextValue::owned(output, reservation)))
}
