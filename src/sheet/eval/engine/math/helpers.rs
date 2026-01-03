use std::result::Result as StdResult;

use crate::sheet::eval::engine::{EvalCtx, flatten_range_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(super) const EPS: f64 = 1e-12;
pub(super) const MAX_BITWISE_VALUE: u64 = (1u64 << 48) - 1;
pub(super) const BIN_MAX: i64 = 511;
pub(super) const BIN_MIN: i64 = -512;
pub(super) const OCT_MAX: i64 = 536_870_911;
pub(super) const OCT_MIN: i64 = -536_870_912;
pub(super) const HEX_MAX: i64 = 549_755_813_887;
pub(super) const HEX_MIN: i64 = -549_755_813_888;

pub(super) fn number_result(value: f64) -> CellValue {
    if value.is_finite()
        && value.fract().abs() < EPS
        && value <= i64::MAX as f64
        && value >= i64::MIN as f64
    {
        CellValue::Int(value as i64)
    } else {
        CellValue::Float(value)
    }
}

pub(super) fn is_even(value: f64) -> bool {
    ((value / 2.0).fract()).abs() < EPS
}

pub(super) fn round_away_from_zero(value: f64) -> f64 {
    let rounded = if value >= 0.0 {
        value.ceil()
    } else {
        value.floor()
    };
    if rounded == -0.0 { 0.0 } else { rounded }
}

pub(super) fn to_int_if_whole(value: f64) -> Option<i64> {
    if !value.is_finite() {
        return None;
    }
    let truncated = value.trunc();
    if (value - truncated).abs() > EPS {
        return None;
    }
    if truncated < i64::MIN as f64 || truncated > i64::MAX as f64 {
        return None;
    }
    Some(truncated as i64)
}

pub(super) async fn flatten_numeric_values(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Vec<f64>> {
    let flat = flatten_range_expr(ctx, current_sheet, expr).await?;
    let values = flat
        .values
        .into_iter()
        .map(|v| to_number(&v).unwrap_or(0.0))
        .collect();
    Ok(values)
}

pub(super) fn to_u48(value: f64) -> Option<u64> {
    if value.is_finite()
        && value >= 0.0
        && value <= MAX_BITWISE_VALUE as f64
        && (value.fract()).abs() < EPS
    {
        Some(value as u64)
    } else {
        None
    }
}

pub(super) fn to_shift_amount(value: f64) -> Option<u32> {
    if value.is_finite() && (0.0..=53.0).contains(&value) && (value.fract()).abs() < EPS {
        Some(value as u32)
    } else {
        None
    }
}

pub(super) fn bit_operand_value(value: &CellValue, func_name: &str) -> StdResult<u64, CellValue> {
    let num = match to_number(value) {
        Some(n) => n,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} arguments must be numeric"
            )));
        },
    };
    match to_u48(num) {
        Some(v) => Ok(v),
        None => Err(CellValue::Error(format!(
            "{func_name} arguments must be integers between 0 and 2^48-1"
        ))),
    }
}

pub(super) fn bit_shift_value(value: &CellValue, func_name: &str) -> StdResult<u32, CellValue> {
    let num = match to_number(value) {
        Some(n) => n,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} shift must be numeric"
            )));
        },
    };
    match to_shift_amount(num) {
        Some(v) => Ok(v),
        None => Err(CellValue::Error(format!(
            "{func_name} shift must be between 0 and 53"
        ))),
    }
}

pub(super) fn binary_string_from_value(
    value: &CellValue,
    func_name: &str,
) -> StdResult<String, CellValue> {
    let raw = match value {
        CellValue::String(s) => s.trim().to_string(),
        CellValue::Int(i) => {
            if *i < 0 {
                return Err(CellValue::Error(format!(
                    "{func_name} expects a binary value containing only 0 or 1"
                )));
            }
            i.to_string()
        },
        CellValue::Float(f) => match to_int_if_whole(*f) {
            Some(i) if i >= 0 => i.to_string(),
            _ => {
                return Err(CellValue::Error(format!(
                    "{func_name} expects a binary value containing only 0 or 1"
                )));
            },
        },
        CellValue::Bool(true) => "1".to_string(),
        CellValue::Bool(false) => "0".to_string(),
        _ => {
            return Err(CellValue::Error(format!(
                "{func_name} expects a binary value containing only 0 or 1"
            )));
        },
    };

    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} expects a binary value containing only 0 or 1"
        )));
    }
    if trimmed.len() > 10 {
        return Err(CellValue::Error(format!(
            "{func_name} expects a binary value up to 10 digits"
        )));
    }
    if !trimmed.chars().all(|c| c == '0' || c == '1') {
        return Err(CellValue::Error(format!(
            "{func_name} expects a binary value containing only 0 or 1"
        )));
    }
    Ok(trimmed.to_string())
}

pub(super) fn parse_signed_binary(bits: &str, func_name: &str) -> StdResult<i64, CellValue> {
    if bits.is_empty() || bits.len() > 10 {
        return Err(CellValue::Error(format!(
            "{func_name} expects a binary value up to 10 digits"
        )));
    }
    let unsigned = u16::from_str_radix(bits, 2).unwrap();
    if bits.len() == 10 && bits.starts_with('1') {
        Ok(unsigned as i64 - 1024)
    } else {
        Ok(unsigned as i64)
    }
}

pub(super) fn parse_places_argument(
    value: &CellValue,
    func_name: &str,
) -> StdResult<usize, CellValue> {
    let num = match to_number(value) {
        Some(n) => n,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} places must be numeric"
            )));
        },
    };
    let int_val = match to_int_if_whole(num) {
        Some(i) => i,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} places must be an integer between 1 and 10"
            )));
        },
    };
    if !(1..=10).contains(&int_val) {
        return Err(CellValue::Error(format!(
            "{func_name} places must be between 1 and 10"
        )));
    }
    Ok(int_val as usize)
}

pub(super) fn pad_with_places(
    result: &mut String,
    places: Option<usize>,
    func_name: &str,
) -> StdResult<(), CellValue> {
    if let Some(p) = places {
        if result.len() > p {
            return Err(CellValue::Error(format!(
                "{func_name} places is too small to display the result"
            )));
        }
        while result.len() < p {
            result.insert(0, '0');
        }
    }
    Ok(())
}

pub(super) fn is_negative_binary(bits: &str) -> bool {
    bits.len() == 10 && bits.starts_with('1')
}

pub(super) fn extend_binary(bits: &str, target_len: usize, fill: char) -> String {
    if bits.len() >= target_len {
        bits.to_string()
    } else {
        let mut extended = String::with_capacity(target_len);
        for _ in 0..(target_len - bits.len()) {
            extended.push(fill);
        }
        extended.push_str(bits);
        extended
    }
}

pub(super) fn negative_binary_to_hex(bits: &str) -> String {
    let extended = extend_binary(bits, 40, '1');
    let value = u64::from_str_radix(&extended, 2).unwrap();
    format!("{value:010X}")
}

pub(super) fn negative_binary_to_oct(bits: &str) -> String {
    let extended = extend_binary(bits, 30, '1');
    let value = u64::from_str_radix(&extended, 2).unwrap();
    format!("{value:010o}")
}

pub(super) fn parse_decimal_for_conversion(
    value: &CellValue,
    func_name: &str,
    min: i64,
    max: i64,
) -> StdResult<i64, CellValue> {
    let num = match to_number(value) {
        Some(n) => n,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} number must be numeric"
            )));
        },
    };
    let int_val = match to_int_if_whole(num) {
        Some(i) => i,
        None => {
            return Err(CellValue::Error(format!(
                "{func_name} number must be an integer"
            )));
        },
    };
    if int_val < min || int_val > max {
        return Err(CellValue::Error(format!(
            "{func_name} number must be between {min} and {max}"
        )));
    }
    Ok(int_val)
}

pub(super) fn ensure_number_in_range(
    value: i64,
    min: i64,
    max: i64,
    func_name: &str,
) -> StdResult<(), CellValue> {
    if value < min || value > max {
        Err(CellValue::Error(format!(
            "{func_name} number must be between {min} and {max}"
        )))
    } else {
        Ok(())
    }
}

pub(super) fn twos_complement_value(number: i64, bits: u32) -> u64 {
    let modulus = 1i128 << bits;
    let adjusted = modulus + number as i128;
    adjusted as u64
}

pub(super) fn parse_hex_string(value: &CellValue, func_name: &str) -> StdResult<String, CellValue> {
    let raw = match value {
        CellValue::String(s) => s.trim().to_string(),
        CellValue::Int(i) => i.to_string(),
        CellValue::Float(f) => match to_int_if_whole(*f) {
            Some(i) => i.to_string(),
            None => {
                return Err(CellValue::Error(format!(
                    "{func_name} expects a hexadecimal value up to 10 digits"
                )));
            },
        },
        CellValue::Bool(true) => "1".to_string(),
        CellValue::Bool(false) => "0".to_string(),
        _ => {
            return Err(CellValue::Error(format!(
                "{func_name} expects a hexadecimal value up to 10 digits"
            )));
        },
    };

    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} expects a hexadecimal value up to 10 digits"
        )));
    }
    let upper = trimmed.to_uppercase();
    let (sign, digits) = if let Some(stripped) = upper.strip_prefix('-') {
        ("-", stripped)
    } else {
        ("", upper.as_str())
    };
    if digits.is_empty() || digits.len() > 10 {
        return Err(CellValue::Error(format!(
            "{func_name} expects a hexadecimal value up to 10 digits"
        )));
    }
    if !digits.chars().all(|c| c.is_ascii_hexdigit()) {
        return Err(CellValue::Error(format!(
            "{func_name} expects a hexadecimal value up to 10 digits"
        )));
    }
    Ok(format!("{sign}{digits}"))
}

pub(super) fn signed_hex_to_decimal(hex: &str, func_name: &str) -> StdResult<i64, CellValue> {
    if let Some(digits) = hex.strip_prefix('-') {
        let value = i64::from_str_radix(digits, 16).unwrap();
        let signed = -value;
        if !(HEX_MIN..=HEX_MAX).contains(&signed) {
            return Err(CellValue::Error(format!(
                "{func_name} number is out of range"
            )));
        }
        return Ok(signed);
    }

    if hex.len() == 10
        && let Some(first) = hex.chars().next()
        && ('8'..='F').contains(&first)
    {
        let raw = u64::from_str_radix(hex, 16).unwrap();
        let signed = raw as i64 - (1i64 << 40);
        return Ok(signed);
    }

    let value = i64::from_str_radix(hex, 16).unwrap();
    if !(HEX_MIN..=HEX_MAX).contains(&value) {
        return Err(CellValue::Error(format!(
            "{func_name} number is out of range"
        )));
    }
    Ok(value)
}

pub(super) fn parse_octal_string(
    value: &CellValue,
    func_name: &str,
) -> StdResult<String, CellValue> {
    let raw = match value {
        CellValue::String(s) => s.trim().to_string(),
        CellValue::Int(i) => i.to_string(),
        CellValue::Float(f) => match to_int_if_whole(*f) {
            Some(i) => i.to_string(),
            None => {
                return Err(CellValue::Error(format!(
                    "{func_name} expects an octal value up to 10 digits"
                )));
            },
        },
        CellValue::Bool(true) => "1".to_string(),
        CellValue::Bool(false) => "0".to_string(),
        _ => {
            return Err(CellValue::Error(format!(
                "{func_name} expects an octal value up to 10 digits"
            )));
        },
    };

    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} expects an octal value up to 10 digits"
        )));
    }
    let (sign, digits) = if let Some(stripped) = trimmed.strip_prefix('-') {
        ("-", stripped)
    } else {
        ("", trimmed)
    };
    if digits.is_empty() || digits.len() > 10 {
        return Err(CellValue::Error(format!(
            "{func_name} expects an octal value up to 10 digits"
        )));
    }
    if !digits.chars().all(|c| ('0'..='7').contains(&c)) {
        return Err(CellValue::Error(format!(
            "{func_name} expects an octal value up to 10 digits"
        )));
    }
    Ok(format!("{sign}{digits}"))
}

pub(super) fn signed_octal_to_decimal(oct: &str, func_name: &str) -> StdResult<i64, CellValue> {
    if let Some(digits) = oct.strip_prefix('-') {
        let value = i64::from_str_radix(digits, 8).unwrap();
        let signed = -value;
        if !(OCT_MIN..=OCT_MAX).contains(&signed) {
            return Err(CellValue::Error(format!(
                "{func_name} number is out of range"
            )));
        }
        return Ok(signed);
    }

    if oct.len() == 10
        && let Some(first) = oct.chars().next()
        && ('4'..='7').contains(&first)
    {
        let raw = u64::from_str_radix(oct, 8).unwrap();
        let signed = raw as i64 - (1i64 << 30);
        return Ok(signed);
    }

    let value = i64::from_str_radix(oct, 8).unwrap();
    if !(OCT_MIN..=OCT_MAX).contains(&value) {
        return Err(CellValue::Error(format!(
            "{func_name} number is out of range"
        )));
    }
    Ok(value)
}

pub(super) fn factorial(n: u64) -> f64 {
    if n <= 1 {
        1.0
    } else {
        (2..=n).fold(1.0, |acc, v| acc * v as f64)
    }
}

pub(super) fn double_factorial(n: u64) -> f64 {
    if n <= 1 {
        1.0
    } else {
        let mut acc = 1.0;
        let mut current = n;
        while current > 1 {
            acc *= current as f64;
            current -= 2;
        }
        acc
    }
}

pub(super) fn combination(n: u64, k: u64) -> f64 {
    if k == 0 || k == n {
        return 1.0;
    }
    let k = k.min(n - k);
    let mut result = 1.0;
    for i in 1..=k {
        result *= (n - k + i) as f64;
        result /= i as f64;
    }
    result
}

pub(super) fn permutation(n: u64, k: u64) -> f64 {
    if k == 0 {
        return 1.0;
    }
    let mut result = 1.0;
    for i in 0..k {
        result *= (n - i) as f64;
    }
    result
}
