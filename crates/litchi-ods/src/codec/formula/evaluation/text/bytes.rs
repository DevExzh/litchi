//! UTF-8 byte-position text functions from OpenFormula §6.7.
//!
//! The byte profile counts octets in the semantic UTF-8 representation while
//! retaining complete Unicode scalars in every returned Text value.  Interior
//! start positions snap backwards to the scalar boundary that contains them.

use super::super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextValue, WorkingValue,
    to_number, to_text,
};

use super::{core, search};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ByteFunction {
    Find,
    Left,
    Len,
    Mid,
    Replace,
    Right,
    Search,
}

impl ByteFunction {
    fn from_name(name: &str) -> Option<Self> {
        Some(match name.len() {
            5 if name.eq_ignore_ascii_case("FINDB") => Self::Find,
            5 if name.eq_ignore_ascii_case("LEFTB") => Self::Left,
            4 if name.eq_ignore_ascii_case("LENB") => Self::Len,
            4 if name.eq_ignore_ascii_case("MIDB") => Self::Mid,
            8 if name.eq_ignore_ascii_case("REPLACEB") => Self::Replace,
            6 if name.eq_ignore_ascii_case("RIGHTB") => Self::Right,
            7 if name.eq_ignore_ascii_case("SEARCHB") => Self::Search,
            _ => return None,
        })
    }

    fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Find | Self::Search => (2..=3).contains(&count),
            Self::Left | Self::Right => (1..=2).contains(&count),
            Self::Len => count == 1,
            Self::Mid => count == 3,
            Self::Replace => count == 4,
        }
    }

    fn text_argument(self, index: usize) -> bool {
        match self {
            Self::Find | Self::Search => index < 2,
            Self::Left | Self::Len | Self::Mid | Self::Right => index == 0,
            Self::Replace => index == 0 || index == 3,
        }
    }
}

pub(super) fn is_byte_function(name: &str) -> bool {
    ByteFunction::from_name(name).is_some()
}

pub(super) fn argument_is_text(name: &str, index: usize) -> bool {
    ByteFunction::from_name(name).is_some_and(|function| function.text_argument(index))
}

pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = ByteFunction::from_name(name).ok_or(EvaluationFailure::Unsupported(
        super::super::UnsupportedKind::Function,
    ))?;
    let count = node.child_count();
    if !function.valid_arity(count) {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = core::first_formula_error(evaluator, node, count)? {
        core::discard_tail(evaluator, count)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if has_required_missing(function, node) {
        core::discard_tail(evaluator, count)?;
        return push_error(evaluator, ScalarError::Value);
    }
    if has_complex_text_argument(evaluator, function, count) {
        core::discard_tail(evaluator, count)?;
        return push_error(evaluator, ScalarError::Value);
    }
    core::reverse_value_tail(evaluator, count)?;

    match function {
        ByteFunction::Find => apply_find(evaluator, node, false),
        ByteFunction::Left => apply_left(evaluator, node),
        ByteFunction::Len => apply_len(evaluator),
        ByteFunction::Mid => apply_mid(evaluator),
        ByteFunction::Replace => apply_replace(evaluator),
        ByteFunction::Right => apply_right(evaluator, node),
        ByteFunction::Search => apply_find(evaluator, node, true),
    }
}

fn has_required_missing(function: ByteFunction, node: Node<'_>) -> bool {
    (0..node.child_count()).any(|index| {
        if !missing(node, index) {
            return false;
        }
        match function {
            ByteFunction::Find | ByteFunction::Search => index != 2,
            ByteFunction::Left | ByteFunction::Right => index != 1,
            _ => true,
        }
    })
}

fn has_complex_text_argument(
    evaluator: &Evaluator<'_, '_, '_>,
    function: ByteFunction,
    count: usize,
) -> bool {
    let start = evaluator.values.len().saturating_sub(count);
    (0..count).any(|index| {
        function.text_argument(index)
            && matches!(
                evaluator.values.get(start + index),
                Some(WorkingValue::Complex(_))
            )
    })
}

fn missing(node: Node<'_>, index: usize) -> bool {
    node.child(index).is_some_and(|child| child.is_missing())
}

fn pop_text<'a>(evaluator: &mut Evaluator<'a, '_, '_>) -> EvaluationResult<TextValue<'a>> {
    to_text(evaluator.pop_value()?, evaluator)
}

fn pop_number(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<Result<f64, ScalarError>> {
    match evaluator.pop_value()? {
        WorkingValue::Text(text) => {
            core::charge_text(evaluator, text.text.as_ref())?;
            // This byte profile distinguishes malformed numeric Text from a
            // parsed non-finite value. The shared integer validators below
            // classify the latter as #NUM rather than losing that distinction
            // in the general scalar Text-to-Number conversion.
            Ok(fast_float2::parse::<f64, _>(text.text.as_ref()).map_err(|_| ScalarError::Value))
        },
        value => to_number(value, evaluator),
    }
}

fn push_error(evaluator: &mut Evaluator<'_, '_, '_>, error: ScalarError) -> EvaluationResult<()> {
    evaluator.push_value(WorkingValue::Error(error))
}

fn push_number(evaluator: &mut Evaluator<'_, '_, '_>, value: f64) -> EvaluationResult<()> {
    if value.is_finite() {
        evaluator.push_value(WorkingValue::Number(value))
    } else {
        push_error(evaluator, ScalarError::Number)
    }
}

fn apply_len(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text(evaluator)?;
    core::charge_text(evaluator, text.text.as_ref())?;
    push_number(evaluator, text.text.len() as f64)
}

fn apply_left<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let text = pop_text(evaluator)?;
    core::charge_text(evaluator, text.text.as_ref())?;
    let length = if node.child_count() < 2 || missing(node, 1) {
        if node.child_count() == 2 {
            let _ = evaluator.pop_value()?;
        }
        1
    } else {
        match core::floor_nonnegative(pop_number(evaluator)?) {
            Ok(value) => value,
            Err(error) => return push_error(evaluator, error),
        }
    };
    let end = left_end(evaluator, text.text.as_ref(), length)?;
    push_text(evaluator, core::text_slice(text, 0, end))
}

fn apply_right<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let text = pop_text(evaluator)?;
    core::charge_text(evaluator, text.text.as_ref())?;
    let length = if node.child_count() < 2 || missing(node, 1) {
        if node.child_count() == 2 {
            let _ = evaluator.pop_value()?;
        }
        1
    } else {
        match core::floor_nonnegative(pop_number(evaluator)?) {
            Ok(value) => value,
            Err(error) => return push_error(evaluator, error),
        }
    };
    let end = text.text.len();
    let start = right_start(evaluator, text.text.as_ref(), length)?;
    push_text(evaluator, core::text_slice(text, start, end))
}

fn apply_mid<'a>(evaluator: &mut Evaluator<'a, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text(evaluator)?;
    let start_result = pop_number(evaluator)?;
    let length_result = pop_number(evaluator)?;
    let start = match core::floor_positive(start_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    let length = match core::floor_nonnegative(length_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    core::charge_text(evaluator, text.text.as_ref())?;
    let Some(start_offset) = normalized_start(evaluator, text.text.as_ref(), start)? else {
        return push_text(evaluator, TextValue::borrowed(""));
    };
    let end = byte_end(evaluator, text.text.as_ref(), start_offset, length)?;
    push_text(evaluator, core::text_slice(text, start_offset, end))
}

fn apply_replace<'a>(evaluator: &mut Evaluator<'a, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text(evaluator)?;
    let start_result = pop_number(evaluator)?;
    let length_result = pop_number(evaluator)?;
    let replacement = pop_text(evaluator)?;
    let start = match core::trunc_positive(start_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    let length = match core::trunc_nonnegative(length_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    core::charge_text(evaluator, text.text.as_ref())?;
    core::charge_text(evaluator, replacement.text.as_ref())?;
    let start = start.min(text.text.len().saturating_add(1));
    let start_offset =
        normalized_start(evaluator, text.text.as_ref(), start)?.unwrap_or(text.text.len());
    let end = byte_end(evaluator, text.text.as_ref(), start_offset, length)?;
    let output_len = start_offset
        .checked_add(replacement.text.len())
        .and_then(|value| value.checked_add(text.text.len().saturating_sub(end)))
        .ok_or_else(|| core::text_limit(evaluator, usize::MAX))?;
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = core::new_output(evaluator, output_len)?;
    output.push_str(&text.text[..start_offset]);
    core::check_execution(evaluator)?;
    output.push_str(replacement.text.as_ref());
    core::check_execution(evaluator)?;
    output.push_str(&text.text[end..]);
    core::check_execution(evaluator)?;
    drop(text);
    drop(replacement);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_find<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'_>,
    insensitive: bool,
) -> EvaluationResult<()> {
    let needle = pop_text(evaluator)?;
    let haystack = pop_text(evaluator)?;
    let start_result = if node.child_count() == 3 {
        if missing(node, 2) {
            let _ = evaluator.pop_value()?;
            None
        } else {
            Some(pop_number(evaluator)?)
        }
    } else {
        None
    };
    let start = match start_result {
        Some(value) => match core::trunc_positive(value) {
            Ok(value) => value,
            Err(error) => return push_error(evaluator, error),
        },
        None => 1,
    };
    core::charge_text(evaluator, needle.text.as_ref())?;
    core::charge_text(evaluator, haystack.text.as_ref())?;
    let Some(start_offset) = normalized_start(evaluator, haystack.text.as_ref(), start)? else {
        return push_error(evaluator, ScalarError::Value);
    };
    if needle.text.is_empty() {
        return push_number(evaluator, start_offset as f64 + 1.0);
    }
    let position = if insensitive {
        let scalar_start =
            scalar_position_for_byte(evaluator, haystack.text.as_ref(), start_offset)?;
        let scalar = search::find(
            evaluator,
            needle.text.as_ref(),
            haystack.text.as_ref(),
            scalar_start,
        )?;
        match scalar {
            Some(scalar) => scalar_to_byte_position(evaluator, haystack.text.as_ref(), scalar)?,
            None => None,
        }
    } else {
        find_literal(
            evaluator,
            needle.text.as_ref(),
            haystack.text.as_ref(),
            start_offset,
        )?
    };
    match position {
        Some(position) => push_number(evaluator, position as f64),
        None => push_error(evaluator, ScalarError::Value),
    }
}

fn find_literal(
    evaluator: &Evaluator<'_, '_, '_>,
    needle: &str,
    haystack: &str,
    start_offset: usize,
) -> EvaluationResult<Option<usize>> {
    if start_offset == haystack.len() {
        return Ok(None);
    }
    core::check_execution(evaluator)?;
    let found = haystack[start_offset..].find(needle);
    core::check_execution(evaluator)?;
    Ok(found.map(|offset| start_offset + offset + 1))
}

fn normalized_start(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    position: usize,
) -> EvaluationResult<Option<usize>> {
    let offset = position.saturating_sub(1);
    if offset > text.len() {
        return Ok(None);
    }
    let mut boundary = offset;
    while boundary > 0 && !text.is_char_boundary(boundary) {
        boundary -= 1;
        core::check_execution(evaluator)?;
    }
    Ok(Some(boundary))
}

// Only folded SEARCH needs a scalar position for the shared matcher. Literal
// FIND and empty queries can use the normalized byte offset directly.
fn scalar_position_for_byte(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    offset: usize,
) -> EvaluationResult<usize> {
    let mut position = 1usize;
    for (index, _) in text[..offset].chars().enumerate() {
        core::checkpoint(evaluator, index)?;
        position += 1;
    }
    Ok(position)
}

fn scalar_to_byte_position(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    scalar: usize,
) -> EvaluationResult<Option<usize>> {
    if scalar == 0 {
        return Ok(None);
    }
    let mut count = 0usize;
    for (index, (offset, _)) in text.char_indices().enumerate() {
        core::checkpoint(evaluator, index)?;
        count = index + 1;
        if count == scalar {
            return Ok(Some(offset + 1));
        }
    }
    if scalar == count + 1 {
        Ok(Some(text.len() + 1))
    } else {
        Ok(None)
    }
}

// A valid UTF-8 scalar occupies at most four bytes, so each boundary
// adjustment below takes at most three steps. Input work is charged before
// these helpers run; no decoded-character buffer or full prefix scan is needed.
fn left_end(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    limit: usize,
) -> EvaluationResult<usize> {
    let mut end = limit.min(text.len());
    while !text.is_char_boundary(end) {
        core::check_execution(evaluator)?;
        end -= 1;
    }
    Ok(end)
}

fn right_start(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    limit: usize,
) -> EvaluationResult<usize> {
    let mut start = text.len().saturating_sub(limit);
    while !text.is_char_boundary(start) {
        core::check_execution(evaluator)?;
        start += 1;
    }
    Ok(start)
}

fn byte_end(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    start: usize,
    limit: usize,
) -> EvaluationResult<usize> {
    left_end(evaluator, text, start.saturating_add(limit))
}

fn push_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: TextValue<'a>,
) -> EvaluationResult<()> {
    evaluator.push_value(WorkingValue::Text(text))
}
