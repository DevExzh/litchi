//! OpenFormula 1.4 section 6.19 Roman-number conversion functions.
//!
//! `ARABIC` is deliberately a small, exact scanner: its right-to-left fold
//! is the rule used by the specification for both canonical and noncanonical
//! Roman text.  `ROMAN` keeps the four format-specific subtraction policies
//! separate from format 4.  The first four formats select the greatest
//! admissible numeral chunk at each step; format 4 uses a bounded signed-coin
//! search.  For each coin, only the two representatives of the required
//! residue class need consideration.  Any coefficient whose magnitude is at
//! least the adjacent radix can be replaced by one next-valued coin and a
//! smaller coefficient, so this finite search contains a minimum-length
//! representation rather than relying on an unproved greedy choice.

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextValue, WorkingValue,
    local_limit, map_execution_error,
};
use litchi_core::Resource;

const ROMAN_VALUES: [u32; 7] = [1, 5, 10, 50, 100, 500, 1000];
const ROMAN_SYMBOLS: [u8; 7] = *b"IVXLCDM";
const ROMAN_RATIOS: [u32; 6] = [5, 2, 5, 2, 5, 2];
const MAX_ROMAN_TEXT_BYTES: usize = 64;
const FORMAT_COUNT: usize = 5;

#[derive(Clone, Copy)]
enum Function {
    Arabic,
    Roman,
}

fn function(name: &str) -> Option<Function> {
    if name.eq_ignore_ascii_case("ARABIC") {
        Some(Function::Arabic)
    } else if name.eq_ignore_ascii_case("ROMAN") {
        Some(Function::Roman)
    } else {
        None
    }
}

pub(super) fn is_roman_function(name: &str) -> bool {
    function(name).is_some()
}

/// Apply a section 6.19 Roman conversion after eager argument evaluation.
/// This family is outlined so its bounded text scans and format search do
/// not enlarge the evaluator's scalar hot loop.
#[inline(never)]
pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    match function(name) {
        Some(Function::Arabic) => apply_arabic(evaluator, node),
        Some(Function::Roman) => apply_roman(evaluator, node),
        None => Err(EvaluationFailure::InvalidExpression(
            "unknown Roman conversion reached evaluator",
        )),
    }
}

fn apply_arabic<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    if node.child_count() != 1 {
        return evaluator.finish_invalid_arity(node);
    }

    let value = evaluator.pop_value()?;
    if let Some(error) = formula_error(&value) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let text = super::to_text(value, evaluator)?;
    let result = arabic_value(evaluator, text.text.as_ref());
    drop(text);
    match result? {
        Ok(value) => evaluator.push_value(WorkingValue::Number(value)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn arabic_value(
    evaluator: &mut Evaluator<'_, '_, '_>,
    text: &str,
) -> EvaluationResult<Result<f64, ScalarError>> {
    evaluator.charge_bytes(text.len())?;

    // The maximum possible input under a 64-bit slice is below 2^74 when
    // every byte is a thousand-valued symbol, so i128 is exact without
    // imposing a narrower numeric cap on accepted noncanonical text.
    let mut total = 0_i128;
    let mut maximum = 0_i128;
    let mut processed = 0usize;
    let mut next_check = 4096usize;
    for byte in text.bytes().rev() {
        if processed >= next_check {
            evaluator
                .context
                .execution
                .check()
                .map_err(map_execution_error)?;
            next_check = processed.saturating_add(4096);
        }
        processed = processed.saturating_add(1);
        let value = match roman_value(byte) {
            Some(value) => i128::from(value),
            None => return Ok(Err(ScalarError::Value)),
        };
        if value < maximum {
            total -= value;
        } else {
            total += value;
            maximum = maximum.max(value);
        }
    }

    let value = total as f64;
    if value.is_finite() {
        Ok(Ok(value))
    } else {
        Ok(Err(ScalarError::Number))
    }
}

fn roman_value(byte: u8) -> Option<u32> {
    match byte {
        b'I' | b'i' => Some(1),
        b'V' | b'v' => Some(5),
        b'X' | b'x' => Some(10),
        b'L' | b'l' => Some(50),
        b'C' | b'c' => Some(100),
        b'D' | b'd' => Some(500),
        b'M' | b'm' => Some(1000),
        _ => None,
    }
}

fn apply_roman<'a>(evaluator: &mut Evaluator<'a, '_, '_>, node: Node<'a>) -> EvaluationResult<()> {
    let count = node.child_count();
    if !(1..=2).contains(&count) {
        return evaluator.finish_invalid_arity(node);
    }

    let format = if count == 2 {
        Some(evaluator.pop_value()?)
    } else {
        None
    };
    let number = evaluator.pop_value()?;

    // Eager function evaluation retains the leftmost formula error.  Check
    // it before converting either argument so errors stay formula values.
    if let Some(error) = formula_error(&number) {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    if let Some(format) = &format {
        if let Some(error) = formula_error(format) {
            return evaluator.push_value(WorkingValue::Error(error));
        }
    }

    let number = match super::to_integer(number, evaluator)? {
        Ok(number) => number,
        Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
    };
    if !number.is_finite() || !(0.0..4000.0).contains(&number) {
        return evaluator.push_value(WorkingValue::Error(ScalarError::Number));
    }
    let number = number as u32;

    let format = match format {
        None => 0,
        Some(WorkingValue::Logical(value)) => {
            if value {
                0
            } else {
                4
            }
        },
        Some(value) => match super::to_integer(value, evaluator)? {
            Ok(value) if value.is_finite() && (0.0..=4.0).contains(&value) => value as usize,
            Ok(_) => return evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
            Err(error) => return evaluator.push_value(WorkingValue::Error(error)),
        },
    };

    if number == 0 {
        // The normative body says zero has an empty representation.  This
        // also makes ARABIC(ROMAN(0)) preserve the stated identity.
        return evaluator.push_value(WorkingValue::Text(TextValue::borrowed("")));
    }

    let (bytes, length) = if format == 4 {
        render_simplified(evaluator, number)?
    } else {
        render_format(evaluator, number, format)?
    };
    render_text(evaluator, &bytes[..length])
}

fn render_format(
    evaluator: &mut Evaluator<'_, '_, '_>,
    mut number: u32,
    format: usize,
) -> EvaluationResult<([u8; MAX_ROMAN_TEXT_BYTES], usize)> {
    debug_assert!(format < FORMAT_COUNT - 1);
    let mut output = [0_u8; MAX_ROMAN_TEXT_BYTES];
    let mut length = 0usize;

    while number != 0 {
        let (value, first, second) = next_chunk(number, format, evaluator)?;
        number -= value;
        append_symbol(&mut output, &mut length, first)?;
        if let Some(second) = second {
            append_symbol(&mut output, &mut length, second)?;
        }
    }
    Ok((output, length))
}

/// Select the greatest admissible single symbol or subtractive pair.  A
/// pair whose value ties a single symbol is ignored, which keeps the output
/// shortest and avoids needless forms such as `VX` for five.
fn next_chunk(
    number: u32,
    format: usize,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<(u32, u8, Option<u8>)> {
    let mut best_value = 0_u32;
    let mut best_first = b'\0';
    let mut best_second = None;

    for index in (0..ROMAN_VALUES.len()).rev() {
        evaluator.charge_work(1)?;
        let value = ROMAN_VALUES[index];
        if value <= number && value > best_value {
            best_value = value;
            best_first = ROMAN_SYMBOLS[index];
            best_second = None;
        }
    }

    for smaller in 0..ROMAN_VALUES.len() - 1 {
        if !subtractor_allowed(smaller, format) {
            continue;
        }
        for larger in (smaller + 1)..ROMAN_VALUES.len() {
            evaluator.charge_work(1)?;
            let small = ROMAN_VALUES[smaller];
            let large = ROMAN_VALUES[larger];
            if format <= 1 && large > small.saturating_mul(10) {
                continue;
            }
            let value = large - small;
            if value <= number && value > best_value {
                best_value = value;
                best_first = ROMAN_SYMBOLS[smaller];
                best_second = Some(ROMAN_SYMBOLS[larger]);
            }
        }
    }

    if best_value == 0 {
        return Err(EvaluationFailure::InvalidExpression(
            "Roman chunk construction made no progress",
        ));
    }
    Ok((best_value, best_first, best_second))
}

fn subtractor_allowed(index: usize, format: usize) -> bool {
    match format {
        0 => matches!(index, 0 | 2 | 4),
        1 => matches!(index, 0..=4),
        2 => matches!(index, 0 | 2 | 3 | 4),
        3 => matches!(index, 0..=4),
        _ => false,
    }
}

fn render_simplified(
    evaluator: &mut Evaluator<'_, '_, '_>,
    number: u32,
) -> EvaluationResult<([u8; MAX_ROMAN_TEXT_BYTES], usize)> {
    // For each adjacent pair of denominations, the coefficient must have a
    // fixed residue modulo the pair's ratio.  The positive residue and its
    // negative carry are the only candidates needed for a minimum.  There
    // are six binary choices, so this search is constant and allocation-free.
    let mut best = [0_i32; ROMAN_VALUES.len()];
    let mut best_length = usize::MAX;

    for mask in 0_u32..(1_u32 << ROMAN_RATIOS.len()) {
        evaluator.charge_work(1)?;
        let mut units = number;
        let mut coefficients = [0_i32; ROMAN_VALUES.len()];
        let mut length = 0usize;

        for index in 0..ROMAN_RATIOS.len() {
            evaluator.charge_work(1)?;
            let ratio = ROMAN_RATIOS[index];
            let remainder = units % ratio;
            let coefficient = if remainder != 0 && (mask & (1_u32 << index)) != 0 {
                remainder as i32 - ratio as i32
            } else {
                remainder as i32
            };
            coefficients[index] = coefficient;
            length = length
                .checked_add(coefficient.unsigned_abs() as usize)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "Roman length overflow",
                ))?;
            let numerator = i64::from(units) - i64::from(coefficient);
            units = u32::try_from(numerator / i64::from(ratio))
                .map_err(|_| EvaluationFailure::InvalidExpression("Roman carry overflow"))?;
        }

        coefficients[ROMAN_VALUES.len() - 1] = i32::try_from(units)
            .map_err(|_| EvaluationFailure::InvalidExpression("Roman high coefficient overflow"))?;
        length = length
            .checked_add(units as usize)
            .ok_or(EvaluationFailure::InvalidExpression(
                "Roman length overflow",
            ))?;
        if length < best_length {
            best_length = length;
            best = coefficients;
        }
    }

    if best_length == usize::MAX {
        return Err(EvaluationFailure::InvalidExpression(
            "Roman simplified search found no representation",
        ));
    }

    let mut output = [0_u8; MAX_ROMAN_TEXT_BYTES];
    let mut length = 0usize;
    // Negative lower-valued coins precede the positive coins.  This order
    // makes each negative symbol see its intended larger symbol to the right.
    for index in 0..ROMAN_VALUES.len() {
        let coefficient = best[index];
        if coefficient < 0 {
            append_repeated(
                &mut output,
                &mut length,
                ROMAN_SYMBOLS[index],
                coefficient.unsigned_abs() as usize,
            )?;
        }
    }
    for index in (0..ROMAN_VALUES.len()).rev() {
        let coefficient = best[index];
        if coefficient > 0 {
            append_repeated(
                &mut output,
                &mut length,
                ROMAN_SYMBOLS[index],
                coefficient as usize,
            )?;
        }
    }
    Ok((output, length))
}

fn append_symbol(
    output: &mut [u8; MAX_ROMAN_TEXT_BYTES],
    length: &mut usize,
    symbol: u8,
) -> EvaluationResult<()> {
    append_repeated(output, length, symbol, 1)
}

fn append_repeated(
    output: &mut [u8; MAX_ROMAN_TEXT_BYTES],
    length: &mut usize,
    symbol: u8,
    count: usize,
) -> EvaluationResult<()> {
    let end = length
        .checked_add(count)
        .ok_or(EvaluationFailure::InvalidExpression(
            "Roman output length overflow",
        ))?;
    if end > output.len() {
        return Err(EvaluationFailure::InvalidExpression(
            "Roman output exceeded fixed construction bound",
        ));
    }
    output[*length..end].fill(symbol);
    *length = end;
    Ok(())
}

fn render_text<'a>(evaluator: &mut Evaluator<'a, '_, '_>, bytes: &[u8]) -> EvaluationResult<()> {
    if bytes.len() > evaluator.limits.max_text_bytes {
        return Err(local_limit(
            Resource::Memory,
            u64::try_from(bytes.len()).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        ));
    }
    evaluator.charge_bytes(bytes.len())?;
    let reservation = evaluator.reserve_storage(bytes.len(), "formula roman text")?;
    let mut output = String::new();
    if !bytes.is_empty() {
        output
            .try_reserve_exact(bytes.len())
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula roman text",
                source,
            })?;
        // Every byte is an ASCII Roman symbol, so the fallible reservation
        // above is the only allocation point for this result.
        for &byte in bytes {
            output.push(byte as char);
        }
    }
    evaluator.push_value(WorkingValue::Text(TextValue::owned(output, reservation)))
}

fn formula_error(value: &WorkingValue<'_>) -> Option<ScalarError> {
    match value {
        WorkingValue::Error(error) => Some(*error),
        _ => None,
    }
}
