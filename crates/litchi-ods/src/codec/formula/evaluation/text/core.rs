//! The scalar implementations for the common OpenFormula text functions.
//!
//! This module deliberately works on the evaluator's value-stack tail.  A
//! second argument vector would duplicate every owned TextValue and would
//! make the storage limit depend on the shape of a function call.  Arguments
//! are reversed in place, inspected for source formula errors, then consumed
//! in source order.

use super::super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextCase, TextValue,
    WorkingValue, concatenate, local_limit, to_integer, to_number, to_text,
};
use litchi_core::Resource;
use std::borrow::Cow;

use super::unicode;

/// The common section 6.20 functions, plus the four scalar code-point
/// functions whose policy is independent of the width tables.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TextFunction {
    Clean,
    Concatenate,
    Exact,
    Find,
    Left,
    Len,
    Lower,
    Mid,
    Proper,
    Replace,
    Rept,
    Right,
    Search,
    Substitute,
    T,
    Trim,
    Upper,
    Char,
    Code,
    Unichar,
    Unicode,
}

impl TextFunction {
    fn from_name(name: &str) -> Option<Self> {
        Some(match name.len() {
            1 if name.eq_ignore_ascii_case("T") => Self::T,
            3 if name.eq_ignore_ascii_case("LEN") => Self::Len,
            4 if name.eq_ignore_ascii_case("CODE") => Self::Code,
            4 if name.eq_ignore_ascii_case("FIND") => Self::Find,
            4 if name.eq_ignore_ascii_case("LEFT") => Self::Left,
            3 if name.eq_ignore_ascii_case("MID") => Self::Mid,
            4 if name.eq_ignore_ascii_case("REPT") => Self::Rept,
            5 if name.eq_ignore_ascii_case("RIGHT") => Self::Right,
            4 if name.eq_ignore_ascii_case("TRIM") => Self::Trim,
            4 if name.eq_ignore_ascii_case("CHAR") => Self::Char,
            5 if name.eq_ignore_ascii_case("CLEAN") => Self::Clean,
            5 if name.eq_ignore_ascii_case("EXACT") => Self::Exact,
            5 if name.eq_ignore_ascii_case("LOWER") => Self::Lower,
            6 if name.eq_ignore_ascii_case("PROPER") => Self::Proper,
            5 if name.eq_ignore_ascii_case("UPPER") => Self::Upper,
            6 if name.eq_ignore_ascii_case("SEARCH") => Self::Search,
            7 if name.eq_ignore_ascii_case("REPLACE") => Self::Replace,
            7 if name.eq_ignore_ascii_case("UNICHAR") => Self::Unichar,
            7 if name.eq_ignore_ascii_case("UNICODE") => Self::Unicode,
            11 if name.eq_ignore_ascii_case("CONCATENATE") => Self::Concatenate,
            10 if name.eq_ignore_ascii_case("SUBSTITUTE") => Self::Substitute,
            _ => return None,
        })
    }

    fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Clean
            | Self::Code
            | Self::Char
            | Self::Len
            | Self::Lower
            | Self::Proper
            | Self::T
            | Self::Trim
            | Self::Upper
            | Self::Unichar
            | Self::Unicode => count == 1,
            Self::Left | Self::Right => (1..=2).contains(&count),
            Self::Concatenate => count > 0,
            Self::Exact | Self::Rept => count == 2,
            Self::Replace => count == 4,
            Self::Find | Self::Search => (2..=3).contains(&count),
            Self::Mid => count == 3,
            Self::Substitute => (3..=4).contains(&count),
        }
    }

    fn text_argument(self, index: usize) -> bool {
        match self {
            Self::Concatenate | Self::Exact => true,
            Self::Clean
            | Self::Code
            | Self::Left
            | Self::Len
            | Self::Lower
            | Self::Mid
            | Self::Proper
            | Self::Rept
            | Self::Right
            | Self::Trim
            | Self::Upper
            | Self::Unicode => index == 0,
            Self::Find | Self::Search => index < 2,
            Self::Replace => index == 0 || index == 3,
            Self::Substitute => index < 3,
            Self::T => false,
            Self::Char | Self::Unichar => false,
        }
    }
}

pub(super) fn is_core_function(name: &str) -> bool {
    TextFunction::from_name(name).is_some()
}

pub(super) fn argument_is_text(name: &str, index: usize) -> bool {
    TextFunction::from_name(name)
        .is_some_and(|function| function == TextFunction::T || function.text_argument(index))
}

pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = TextFunction::from_name(name).ok_or(EvaluationFailure::Unsupported(
        super::super::UnsupportedKind::Function,
    ))?;
    let count = node.child_count();
    if !function.valid_arity(count) {
        return evaluator.finish_invalid_arity(node);
    }
    if let Some(error) = first_formula_error(evaluator, node, count)? {
        discard_tail(evaluator, count)?;
        return evaluator.push_value(WorkingValue::Error(error));
    }
    reverse_value_tail(evaluator, count)?;
    if has_required_missing(function, node) {
        discard_tail(evaluator, count)?;
        return push_error(evaluator, ScalarError::Value);
    }
    if has_complex_text_argument(evaluator, function, count) {
        discard_tail(evaluator, count)?;
        return push_error(evaluator, ScalarError::Value);
    }

    match function {
        TextFunction::Clean => apply_clean(evaluator),
        TextFunction::Concatenate => apply_concatenate(evaluator, count),
        TextFunction::Exact => apply_exact(evaluator),
        TextFunction::Find => apply_find(evaluator, node, false),
        TextFunction::Left => apply_left(evaluator, node),
        TextFunction::Len => apply_len(evaluator),
        TextFunction::Lower => apply_case(evaluator, false),
        TextFunction::Mid => apply_mid(evaluator),
        TextFunction::Proper => apply_proper(evaluator),
        TextFunction::Replace => apply_replace(evaluator),
        TextFunction::Rept => apply_rept(evaluator),
        TextFunction::Right => apply_right(evaluator, node),
        TextFunction::Search => apply_find(evaluator, node, true),
        TextFunction::Substitute => apply_substitute(evaluator, node),
        TextFunction::T => apply_t(evaluator, node),
        TextFunction::Trim => apply_trim(evaluator),
        TextFunction::Upper => apply_case(evaluator, true),
        TextFunction::Char => apply_char(evaluator),
        TextFunction::Code => apply_code(evaluator),
        TextFunction::Unichar => apply_unichar(evaluator),
        TextFunction::Unicode => apply_unicode(evaluator),
    }
}

/// Reverse the eagerly scheduled argument tail without allocating another
/// value vector.
pub(super) fn reverse_value_tail<'a>(
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
                "text value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

pub(super) fn discard_tail(
    evaluator: &mut Evaluator<'_, '_, '_>,
    count: usize,
) -> EvaluationResult<()> {
    for _ in 0..count {
        let _ = evaluator.pop_value()?;
    }
    Ok(())
}

/// Find the first source formula error while the tail is in source order.
/// Missing AST slots are generated values and therefore are intentionally
/// skipped here; optional handlers supply their defaults below.
pub(super) fn first_formula_error(
    evaluator: &mut Evaluator<'_, '_, '_>,
    node: Node<'_>,
    count: usize,
) -> EvaluationResult<Option<ScalarError>> {
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "text value stack underflow",
            ))?;
    let mut next_check = 0usize;
    for index in 0..count {
        if index >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(super::super::map_execution_error)?;
            next_check = index.saturating_add(4096);
        }
        if node.child(index).is_some_and(|child| child.is_missing()) {
            continue;
        }
        let physical = start + index;
        if let WorkingValue::Error(error) = &evaluator.values[physical] {
            return Ok(Some(*error));
        }
    }
    Ok(None)
}

fn has_complex_text_argument(
    evaluator: &Evaluator<'_, '_, '_>,
    function: TextFunction,
    count: usize,
) -> bool {
    let start = evaluator.values.len().saturating_sub(count);
    (0..count).any(|index| {
        function.text_argument(index)
            && matches!(
                evaluator.values.get(start + count - 1 - index),
                Some(WorkingValue::Complex(_))
            )
    })
}

pub(super) fn has_complex_argument(evaluator: &Evaluator<'_, '_, '_>, count: usize) -> bool {
    let start = evaluator.values.len().saturating_sub(count);
    (0..count).any(|index| {
        matches!(
            evaluator.values.get(start + count - 1 - index),
            Some(WorkingValue::Complex(_))
        )
    })
}

/// The amount of source text consumed by a scan is charged before the scan.
pub(super) fn charge_text(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
) -> EvaluationResult<()> {
    evaluator.charge_bytes(text.len())
}

/// Check cancellation at a bounded interval in scans whose byte work was
/// charged up front.  The evaluator's regular `charge_work` calls already
/// check before every charged operation; this helper covers copy/filter
/// loops that intentionally charge their complete input or output once.
pub(super) fn checkpoint(evaluator: &Evaluator<'_, '_, '_>, index: usize) -> EvaluationResult<()> {
    if index & 4095 == 0 {
        check_execution(evaluator)?;
    }
    Ok(())
}

pub(super) fn check_execution(evaluator: &Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    evaluator
        .context
        .execution
        .check()
        .map_err(super::super::map_execution_error)
}

pub(super) fn text_limit(evaluator: &Evaluator<'_, '_, '_>, observed: usize) -> EvaluationFailure {
    local_limit(
        Resource::Memory,
        u64::try_from(observed).unwrap_or(u64::MAX),
        u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
    )
}

pub(super) fn new_output(
    evaluator: &Evaluator<'_, '_, '_>,
    length: usize,
) -> EvaluationResult<(litchi_core::Reservation, String)> {
    if length > evaluator.limits.max_text_bytes {
        return Err(text_limit(evaluator, length));
    }
    let reservation = evaluator.reserve_storage(length, "formula scalar text result")?;
    let mut output = String::new();
    if length != 0 {
        output
            .try_reserve_exact(length)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula scalar text result",
                source,
            })?;
    }
    Ok((reservation, output))
}

fn finish_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: TextValue<'a>,
) -> EvaluationResult<()> {
    evaluator.push_value(WorkingValue::Text(text))
}

fn missing(node: Node<'_>, index: usize) -> bool {
    node.child(index).is_some_and(|child| child.is_missing())
}

fn has_required_missing(function: TextFunction, node: Node<'_>) -> bool {
    (0..node.child_count()).any(|index| {
        if !missing(node, index) {
            return false;
        }
        match function {
            TextFunction::Left | TextFunction::Right => index != 1,
            TextFunction::Find | TextFunction::Search => index != 2,
            TextFunction::Substitute => index != 3,
            _ => true,
        }
    })
}

fn pop_text_required<'a>(evaluator: &mut Evaluator<'a, '_, '_>) -> EvaluationResult<TextValue<'a>> {
    let value = evaluator.pop_value()?;
    to_text(value, evaluator)
}

fn pop_number(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<Result<f64, ScalarError>> {
    to_number(evaluator.pop_value()?, evaluator)
}

fn pop_integer(
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    to_integer(evaluator.pop_value()?, evaluator)
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

fn push_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: TextValue<'a>,
) -> EvaluationResult<()> {
    finish_text(evaluator, value)
}

fn is_clean_removed(character: char) -> bool {
    unicode::is_clean_removed(character)
}

pub(super) fn text_slice<'a>(value: TextValue<'a>, start: usize, end: usize) -> TextValue<'a> {
    let length = end.saturating_sub(start);
    match value.text {
        Cow::Borrowed(text) => TextValue::borrowed(&text[start..end]),
        Cow::Owned(mut text) => {
            if start == 0 && end == text.len() {
                return TextValue {
                    text: Cow::Owned(text),
                    reservation: value.reservation,
                };
            }
            if start != 0 {
                drop(text.drain(..start));
            }
            text.truncate(length);
            if text.is_empty() {
                drop(text);
                TextValue::borrowed("")
            } else {
                TextValue {
                    text: Cow::Owned(text),
                    reservation: value.reservation,
                }
            }
        },
    }
}

fn char_boundary(
    evaluator: &Evaluator<'_, '_, '_>,
    text: &str,
    index: usize,
) -> EvaluationResult<usize> {
    for (position, (offset, _)) in text.char_indices().enumerate() {
        checkpoint(evaluator, position)?;
        if position == index {
            return Ok(offset);
        }
    }
    Ok(text.len())
}

fn char_count(evaluator: &Evaluator<'_, '_, '_>, text: &str) -> EvaluationResult<usize> {
    let mut count = 0usize;
    for (position, (_, _)) in text.char_indices().enumerate() {
        checkpoint(evaluator, position)?;
        count = count.saturating_add(1);
    }
    Ok(count)
}

pub(super) fn floor_nonnegative(value: Result<f64, ScalarError>) -> Result<usize, ScalarError> {
    let value = value?;
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    if value < 0.0 {
        return Err(ScalarError::Value);
    }
    let value = value.floor();
    Ok(if value > usize::MAX as f64 {
        usize::MAX
    } else {
        value as usize
    })
}

pub(super) fn floor_positive(value: Result<f64, ScalarError>) -> Result<usize, ScalarError> {
    let value = value?;
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    if value < 1.0 {
        return Err(ScalarError::Value);
    }
    let value = value.floor();
    Ok(if value > usize::MAX as f64 {
        usize::MAX
    } else {
        value as usize
    })
}

pub(super) fn trunc_nonnegative(value: Result<f64, ScalarError>) -> Result<usize, ScalarError> {
    let value = value?;
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    if value < 0.0 {
        return Err(ScalarError::Value);
    }
    let value = value.trunc();
    Ok(if value > usize::MAX as f64 {
        usize::MAX
    } else {
        value as usize
    })
}

pub(super) fn trunc_positive(value: Result<f64, ScalarError>) -> Result<usize, ScalarError> {
    let value = value?;
    if !value.is_finite() {
        return Err(ScalarError::Number);
    }
    if value < 1.0 {
        return Err(ScalarError::Value);
    }
    let value = value.trunc();
    Ok(if value > usize::MAX as f64 {
        usize::MAX
    } else {
        value as usize
    })
}

fn apply_clean(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let source = text.text.as_ref();
    charge_text(evaluator, source)?;
    let mut output_len = 0usize;
    let mut changed = false;
    for (index, (_, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        if is_clean_removed(character) {
            changed = true;
        } else {
            output_len = output_len
                .checked_add(character.len_utf8())
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        }
    }
    if !changed {
        return push_text(evaluator, text);
    }
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    for (index, (_, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        if !is_clean_removed(character) {
            output.push(character);
        }
    }
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_concatenate<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    count: usize,
) -> EvaluationResult<()> {
    let mut result: Option<WorkingValue<'a>> = None;
    for _ in 0..count {
        let value = evaluator.pop_value()?;
        if let WorkingValue::Text(text) = &value {
            charge_text(evaluator, text.text.as_ref())?;
        }
        let value = to_text(value, evaluator)?;
        result = Some(match result.take() {
            Some(left) => {
                let result = concatenate(evaluator, left, WorkingValue::Text(value))?;
                check_execution(evaluator)?;
                result
            },
            None => WorkingValue::Text(value),
        });
    }
    evaluator.push_value(result.ok_or(EvaluationFailure::InvalidExpression(
        "CONCATENATE had no arguments",
    ))?)
}

fn apply_exact(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let left = pop_text_required(evaluator)?;
    let right = pop_text_required(evaluator)?;
    charge_text(evaluator, left.text.as_ref())?;
    charge_text(evaluator, right.text.as_ref())?;
    let equal = super::super::compare_text(
        evaluator,
        left.text.as_ref(),
        right.text.as_ref(),
        TextCase::Sensitive,
    )? == std::cmp::Ordering::Equal;
    evaluator.push_value(WorkingValue::Logical(equal))
}

fn apply_find(
    evaluator: &mut Evaluator<'_, '_, '_>,
    node: Node<'_>,
    insensitive: bool,
) -> EvaluationResult<()> {
    let needle = pop_text_required(evaluator)?;
    let haystack = pop_text_required(evaluator)?;
    charge_text(evaluator, needle.text.as_ref())?;
    charge_text(evaluator, haystack.text.as_ref())?;
    let start = if node.child_count() == 3 {
        if missing(node, 2) {
            let _ = evaluator.pop_value()?;
            1usize
        } else {
            match trunc_positive(pop_number(evaluator)?) {
                Ok(value) => value,
                Err(error) => return push_error(evaluator, error),
            }
        }
    } else {
        1
    };
    let position = find_position(
        evaluator,
        needle.text.as_ref(),
        haystack.text.as_ref(),
        start,
        insensitive,
    )?;
    match position {
        Some(position) => push_number(evaluator, position as f64),
        None => push_error(evaluator, ScalarError::Value),
    }
}

fn find_position(
    evaluator: &mut Evaluator<'_, '_, '_>,
    needle: &str,
    haystack: &str,
    start: usize,
    insensitive: bool,
) -> EvaluationResult<Option<usize>> {
    let start_index = start.saturating_sub(1);
    let length = char_count(evaluator, haystack)?;
    if start_index > length {
        return Ok(None);
    }
    if needle.is_empty() {
        return Ok(Some(start));
    }
    let start_byte = char_boundary(evaluator, haystack, start_index)?;
    if !insensitive {
        check_execution(evaluator)?;
        let found = haystack[start_byte..].find(needle);
        check_execution(evaluator)?;
        let Some(offset) = found else {
            return Ok(None);
        };
        let skipped = char_count(evaluator, &haystack[start_byte..start_byte + offset])?;
        return Ok(Some(start_index + skipped + 1));
    }

    super::search::find(evaluator, needle, haystack, start)
}

fn apply_left(evaluator: &mut Evaluator<'_, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    charge_text(evaluator, text.text.as_ref())?;
    let length = if node.child_count() < 2 || missing(node, 1) {
        if node.child_count() >= 2 {
            let _ = evaluator.pop_value()?;
        }
        1
    } else {
        match floor_nonnegative(pop_number(evaluator)?) {
            Ok(value) => value,
            Err(error) => return push_error(evaluator, error),
        }
    };
    let characters = char_count(evaluator, text.text.as_ref())?;
    let end = char_boundary(evaluator, text.text.as_ref(), length.min(characters))?;
    push_text(evaluator, text_slice(text, 0, end))
}

fn apply_right(evaluator: &mut Evaluator<'_, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    charge_text(evaluator, text.text.as_ref())?;
    let length = if node.child_count() < 2 || missing(node, 1) {
        if node.child_count() >= 2 {
            let _ = evaluator.pop_value()?;
        }
        1
    } else {
        match floor_nonnegative(pop_number(evaluator)?) {
            Ok(value) => value,
            Err(error) => return push_error(evaluator, error),
        }
    };
    let source = text.text.as_ref();
    let characters = char_count(evaluator, source)?;
    let length = length.min(characters);
    let start = char_boundary(evaluator, source, characters - length)?;
    let end = source.len();
    push_text(evaluator, text_slice(text, start, end))
}

fn apply_len(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    charge_text(evaluator, text.text.as_ref())?;
    let count = char_count(evaluator, text.text.as_ref())?;
    push_number(evaluator, count as f64)
}

fn apply_mid(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let start_result = pop_number(evaluator)?;
    let length_result = pop_number(evaluator)?;
    let start = match floor_positive(start_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    let length = match floor_nonnegative(length_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    charge_text(evaluator, text.text.as_ref())?;
    let source = text.text.as_ref();
    let characters = char_count(evaluator, source)?;
    if start > characters || length == 0 {
        return push_text(evaluator, TextValue::borrowed(""));
    }
    let start_byte = char_boundary(evaluator, source, start - 1)?;
    let end_byte = char_boundary(
        evaluator,
        source,
        (start - 1).saturating_add(length).min(characters),
    )?;
    push_text(evaluator, text_slice(text, start_byte, end_byte))
}

fn is_cased(character: char) -> bool {
    unicode::is_cased(character)
}

fn is_case_ignorable(character: char) -> bool {
    unicode::is_case_ignorable(character)
}

fn final_sigma(
    evaluator: &Evaluator<'_, '_, '_>,
    source: &str,
    offset: usize,
) -> EvaluationResult<bool> {
    let before = &source[..offset];
    let after = &source[offset + '\u{03a3}'.len_utf8()..];
    let mut preceded = false;
    for (index, character) in before.chars().rev().enumerate() {
        checkpoint(evaluator, index)?;
        if !is_case_ignorable(character) {
            preceded = is_cased(character);
            break;
        }
    }
    let mut followed = false;
    for (index, character) in after.chars().enumerate() {
        checkpoint(evaluator, index)?;
        if !is_case_ignorable(character) {
            followed = is_cased(character);
            break;
        }
    }
    Ok(preceded && !followed)
}

fn write_lowercase(
    evaluator: &Evaluator<'_, '_, '_>,
    output: &mut String,
    character: char,
    source: &str,
    offset: usize,
) -> EvaluationResult<()> {
    if character == '\u{03a3}' && final_sigma(evaluator, source, offset)? {
        output.push('\u{03c2}');
    } else {
        output.extend(unicode::lowercase(character).iter());
    }
    Ok(())
}

fn lowercase_len(
    evaluator: &Evaluator<'_, '_, '_>,
    character: char,
    source: &str,
    offset: usize,
) -> EvaluationResult<usize> {
    if character == '\u{03a3}' && final_sigma(evaluator, source, offset)? {
        Ok('ς'.len_utf8())
    } else {
        Ok(unicode::lowercase(character)
            .iter()
            .map(char::len_utf8)
            .sum())
    }
}

fn lowercase_changes(
    evaluator: &Evaluator<'_, '_, '_>,
    character: char,
    source: &str,
    offset: usize,
) -> EvaluationResult<bool> {
    if character == '\u{03a3}' && final_sigma(evaluator, source, offset)? {
        return Ok(character != '\u{03c2}');
    }
    let mut mapping = unicode::lowercase(character).iter();
    Ok(match mapping.next() {
        Some(first) if first != character => true,
        Some(_) => mapping.any(|mapped| mapped != character),
        None => true,
    })
}

fn uppercase_changes(character: char) -> bool {
    let mut mapping = unicode::uppercase(character).iter();
    match mapping.next() {
        Some(first) if first != character => true,
        Some(_) => mapping.any(|mapped| mapped != character),
        None => true,
    }
}

fn apply_case(evaluator: &mut Evaluator<'_, '_, '_>, upper: bool) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let source = text.text.as_ref();
    charge_text(evaluator, source)?;
    let mut output_len = 0usize;
    let mut changed = false;
    for (index, (offset, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        let length = if upper {
            unicode::uppercase(character)
                .iter()
                .map(char::len_utf8)
                .sum()
        } else {
            lowercase_len(evaluator, character, source, offset)?
        };
        output_len = output_len
            .checked_add(length)
            .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        if length != character.len_utf8()
            || if upper {
                uppercase_changes(character)
            } else {
                lowercase_changes(evaluator, character, source, offset)?
            }
        {
            changed = true;
        }
    }
    if !changed {
        return push_text(evaluator, text);
    }
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    for (index, (offset, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        if upper {
            output.extend(unicode::uppercase(character).iter());
        } else {
            write_lowercase(evaluator, &mut output, character, source, offset)?;
        }
    }
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_proper(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let source = text.text.as_ref();
    charge_text(evaluator, source)?;
    let mut output_len = 0usize;
    let mut changed = false;
    let mut previous_letter = false;
    for (index, (offset, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        let letter = unicode::is_letter(character);
        let upper = letter && !previous_letter;
        let length = if upper {
            unicode::uppercase(character)
                .iter()
                .map(char::len_utf8)
                .sum()
        } else if letter {
            lowercase_len(evaluator, character, source, offset)?
        } else {
            character.len_utf8()
        };
        output_len = output_len
            .checked_add(length)
            .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        let different = if upper {
            uppercase_changes(character)
        } else if letter {
            lowercase_changes(evaluator, character, source, offset)?
        } else {
            false
        };
        if different {
            changed = true;
        }
        previous_letter = letter;
    }
    if !changed {
        return push_text(evaluator, text);
    }
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    previous_letter = false;
    for (index, (offset, character)) in source.char_indices().enumerate() {
        checkpoint(evaluator, index)?;
        let letter = unicode::is_letter(character);
        if letter && !previous_letter {
            output.extend(unicode::uppercase(character).iter());
        } else if letter {
            write_lowercase(evaluator, &mut output, character, source, offset)?;
        } else {
            output.push(character);
        }
        previous_letter = letter;
    }
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_replace(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let start_result = pop_number(evaluator)?;
    let count_result = pop_number(evaluator)?;
    let replacement = pop_text_required(evaluator)?;
    let start = match trunc_positive(start_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    let count = match trunc_nonnegative(count_result) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    charge_text(evaluator, text.text.as_ref())?;
    charge_text(evaluator, replacement.text.as_ref())?;
    let source = text.text.as_ref();
    let characters = char_count(evaluator, source)?;
    // Follow the coherent LEFT/MID equation: a start beyond the source
    // appends the replacement. Older copies of the prose say to clamp it to
    // the last letter, which contradicts that equation.
    let start_index = if start > characters {
        characters
    } else {
        start.saturating_sub(1)
    };
    let start_byte = char_boundary(evaluator, source, start_index)?;
    let end_byte = char_boundary(
        evaluator,
        source,
        start_index.saturating_add(count).min(characters),
    )?;
    let prefix_len = start_byte;
    let suffix_len = source.len().saturating_sub(end_byte);
    let output_len = prefix_len
        .checked_add(replacement.text.len())
        .and_then(|length| length.checked_add(suffix_len))
        .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    output.push_str(&source[..start_byte]);
    check_execution(evaluator)?;
    output.push_str(replacement.text.as_ref());
    check_execution(evaluator)?;
    output.push_str(&source[end_byte..]);
    check_execution(evaluator)?;
    drop(text);
    drop(replacement);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_rept(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let count = match trunc_nonnegative(pop_number(evaluator)?) {
        Ok(value) => value,
        Err(error) => return push_error(evaluator, error),
    };
    let source = text.text.as_ref();
    charge_text(evaluator, source)?;
    if count == 0 || source.is_empty() {
        drop(text);
        return push_text(evaluator, TextValue::borrowed(""));
    }
    if count == 1 {
        return push_text(evaluator, text);
    }
    let output_len = source
        .len()
        .checked_mul(count)
        .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    for repetition in 0..count {
        checkpoint(evaluator, repetition)?;
        output.push_str(source);
        check_execution(evaluator)?;
    }
    drop(text);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_substitute(evaluator: &mut Evaluator<'_, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let old = pop_text_required(evaluator)?;
    let new = pop_text_required(evaluator)?;
    let which = if node.child_count() == 4 {
        if missing(node, 3) {
            let _ = evaluator.pop_value()?;
            None
        } else {
            match trunc_positive(pop_number(evaluator)?) {
                Ok(value) => Some(value),
                Err(error) => return push_error(evaluator, error),
            }
        }
    } else {
        None
    };
    let source = text.text.as_ref();
    let old_text = old.text.as_ref();
    let new_text = new.text.as_ref();
    charge_text(evaluator, source)?;
    charge_text(evaluator, old_text)?;
    charge_text(evaluator, new_text)?;
    if old_text.is_empty() {
        drop(old);
        drop(new);
        return push_text(evaluator, text);
    }

    let mut occurrences = 0usize;
    let mut cursor = 0usize;
    let mut output_len = source.len();
    let mut search_iteration = 0usize;
    while let Some(relative) = {
        checkpoint(evaluator, search_iteration)?;
        let found = source[cursor..].find(old_text);
        check_execution(evaluator)?;
        search_iteration = search_iteration.saturating_add(1);
        found
    } {
        let start = cursor + relative;
        occurrences = occurrences
            .checked_add(1)
            .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        let selected = which.is_none_or(|value| value == occurrences);
        if selected {
            output_len = output_len
                .checked_sub(old_text.len())
                .and_then(|length| length.checked_add(new_text.len()))
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        }
        cursor = start
            .checked_add(old_text.len())
            .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        if cursor > source.len() {
            break;
        }
    }
    if occurrences == 0 || which.is_some_and(|value| value > occurrences) {
        drop(old);
        drop(new);
        return push_text(evaluator, text);
    }

    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    cursor = 0;
    let mut occurrence = 0usize;
    search_iteration = 0;
    while let Some(relative) = {
        checkpoint(evaluator, search_iteration)?;
        let found = source[cursor..].find(old_text);
        check_execution(evaluator)?;
        search_iteration = search_iteration.saturating_add(1);
        found
    } {
        let start = cursor + relative;
        output.push_str(&source[cursor..start]);
        check_execution(evaluator)?;
        occurrence += 1;
        if which.is_none_or(|value| value == occurrence) {
            output.push_str(new_text);
        } else {
            output.push_str(old_text);
        }
        check_execution(evaluator)?;
        cursor = start + old_text.len();
        if cursor > source.len() {
            break;
        }
        if which.is_some_and(|value| occurrence >= value) {
            break;
        }
    }
    output.push_str(&source[cursor..]);
    check_execution(evaluator)?;
    drop(text);
    drop(old);
    drop(new);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn is_trim_space(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\n' | '\r')
}

fn trimmed_shape(
    evaluator: &Evaluator<'_, '_, '_>,
    source: &str,
) -> Result<(usize, bool), EvaluationFailure> {
    let mut output_len = 0usize;
    let mut changed = false;
    let mut started = false;
    let mut iterator = source.chars().peekable();
    let mut position = 0usize;
    while let Some(character) = iterator.next() {
        checkpoint(evaluator, position)?;
        position = position.saturating_add(1);
        if !is_trim_space(character) {
            output_len = output_len
                .checked_add(character.len_utf8())
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
            started = true;
            continue;
        }
        let mut run_length = 1usize;
        while iterator.peek().copied().is_some_and(is_trim_space) {
            iterator.next();
            checkpoint(evaluator, position)?;
            position = position.saturating_add(1);
            run_length = run_length
                .checked_add(1)
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        }
        let internal = started && iterator.peek().is_some();
        if !internal {
            changed = true;
        } else if run_length == 1 {
            output_len = output_len
                .checked_add(character.len_utf8())
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
        } else {
            output_len = output_len
                .checked_add(1)
                .ok_or_else(|| text_limit(evaluator, usize::MAX))?;
            changed = true;
        }
    }
    Ok((output_len, changed))
}

fn apply_trim(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    let source = text.text.as_ref();
    charge_text(evaluator, source)?;
    let (output_len, changed) = trimmed_shape(evaluator, source)?;
    if !changed {
        return push_text(evaluator, text);
    }
    evaluator.charge_bytes(output_len)?;
    let (reservation, mut output) = new_output(evaluator, output_len)?;
    let mut started = false;
    let mut iterator = source.chars().peekable();
    let mut position = 0usize;
    while let Some(character) = iterator.next() {
        checkpoint(evaluator, position)?;
        position = position.saturating_add(1);
        if is_trim_space(character) {
            let mut run_length = 1usize;
            while iterator.peek().copied().is_some_and(is_trim_space) {
                iterator.next();
                checkpoint(evaluator, position)?;
                position = position.saturating_add(1);
                run_length += 1;
            }
            if started && iterator.peek().is_some() {
                if run_length == 1 {
                    output.push(character);
                } else {
                    output.push(' ');
                }
            }
            continue;
        }
        output.push(character);
        started = true;
    }
    drop(text);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_t(evaluator: &mut Evaluator<'_, '_, '_>, node: Node<'_>) -> EvaluationResult<()> {
    let value = evaluator.pop_value()?;
    if missing(node, 0) {
        return push_text(evaluator, TextValue::borrowed(""));
    }
    match value {
        WorkingValue::Text(text) => {
            charge_text(evaluator, text.text.as_ref())?;
            push_text(evaluator, text)
        },
        WorkingValue::Error(error) => push_error(evaluator, error),
        _ => push_text(evaluator, TextValue::borrowed("")),
    }
}

fn apply_char(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let value = match pop_integer(evaluator)? {
        Ok(value) if value.is_finite() && (1.0..=255.0).contains(&value) => value as u32,
        Ok(value) if value.is_finite() => {
            return push_error(evaluator, ScalarError::Value);
        },
        Ok(_) => return push_error(evaluator, ScalarError::Number),
        Err(error) => return push_error(evaluator, error),
    };
    let character = char::from_u32(value).ok_or(EvaluationFailure::InvalidExpression(
        "CHAR accepted a non-Unicode scalar",
    ))?;
    evaluator.charge_bytes(character.len_utf8())?;
    let (reservation, mut output) = new_output(evaluator, character.len_utf8())?;
    output.push(character);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_code(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    charge_text(evaluator, text.text.as_ref())?;
    match text.text.chars().next() {
        Some(character) => push_number(evaluator, character as u32 as f64),
        None => push_error(evaluator, ScalarError::Value),
    }
}

fn apply_unichar(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let value = match pop_integer(evaluator)? {
        Ok(value) if value.is_finite() && (0.0..=0x10ffff as f64).contains(&value) => value as u32,
        Ok(_) => return push_error(evaluator, ScalarError::Value),
        Err(error) => return push_error(evaluator, error),
    };
    let Some(character) = char::from_u32(value) else {
        return push_error(evaluator, ScalarError::Value);
    };
    evaluator.charge_bytes(character.len_utf8())?;
    let (reservation, mut output) = new_output(evaluator, character.len_utf8())?;
    output.push(character);
    push_text(evaluator, TextValue::owned(output, reservation))
}

fn apply_unicode(evaluator: &mut Evaluator<'_, '_, '_>) -> EvaluationResult<()> {
    let text = pop_text_required(evaluator)?;
    charge_text(evaluator, text.text.as_ref())?;
    match text.text.chars().next() {
        Some(character) => push_number(evaluator, character as u32 as f64),
        None => push_error(evaluator, ScalarError::Value),
    }
}
