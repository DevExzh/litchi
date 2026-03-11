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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_result_int() {
        let result = number_result(5.0);
        assert!(matches!(result, CellValue::Int(5)));
    }

    #[test]
    fn test_number_result_float() {
        let result = number_result(5.5);
        assert!(matches!(result, CellValue::Float(v) if (v - 5.5).abs() < 1e-9));
    }

    #[test]
    fn test_number_result_large_float() {
        let result = number_result(1e20);
        assert!(matches!(result, CellValue::Float(_)));
    }

    #[test]
    fn test_is_even_true() {
        assert!(is_even(4.0));
        assert!(is_even(-2.0));
        assert!(is_even(0.0));
    }

    #[test]
    fn test_is_even_false() {
        assert!(!is_even(3.0));
        assert!(!is_even(-1.0));
    }

    #[test]
    fn test_round_away_from_zero() {
        assert_eq!(round_away_from_zero(2.3), 3.0);
        assert_eq!(round_away_from_zero(2.7), 3.0);
        assert_eq!(round_away_from_zero(-2.3), -3.0);
        assert_eq!(round_away_from_zero(-2.7), -3.0);
        assert_eq!(round_away_from_zero(0.0), 0.0);
    }

    #[test]
    fn test_to_int_if_whole() {
        assert_eq!(to_int_if_whole(5.0), Some(5));
        assert_eq!(to_int_if_whole(-5.0), Some(-5));
        assert_eq!(to_int_if_whole(5.5), None);
        assert_eq!(to_int_if_whole(f64::NAN), None);
        assert_eq!(to_int_if_whole(f64::INFINITY), None);
    }

    #[test]
    fn test_to_u48_valid() {
        assert_eq!(to_u48(0.0), Some(0));
        assert_eq!(to_u48(100.0), Some(100));
        assert_eq!(to_u48(((1u64 << 48) - 1) as f64), Some((1u64 << 48) - 1));
    }

    #[test]
    fn test_to_u48_invalid() {
        assert_eq!(to_u48(-1.0), None);
        assert_eq!(to_u48((1u64 << 48) as f64), None);
        assert_eq!(to_u48(5.5), None);
        assert_eq!(to_u48(f64::NAN), None);
    }

    #[test]
    fn test_to_shift_amount_valid() {
        assert_eq!(to_shift_amount(0.0), Some(0));
        assert_eq!(to_shift_amount(10.0), Some(10));
        assert_eq!(to_shift_amount(53.0), Some(53));
    }

    #[test]
    fn test_to_shift_amount_invalid() {
        assert_eq!(to_shift_amount(-1.0), None);
        assert_eq!(to_shift_amount(54.0), None);
        assert_eq!(to_shift_amount(5.5), None);
    }

    #[test]
    fn test_factorial() {
        assert_eq!(factorial(0), 1.0);
        assert_eq!(factorial(1), 1.0);
        assert_eq!(factorial(5), 120.0);
        assert_eq!(factorial(10), 3628800.0);
    }

    #[test]
    fn test_double_factorial() {
        assert_eq!(double_factorial(0), 1.0);
        assert_eq!(double_factorial(1), 1.0);
        assert_eq!(double_factorial(5), 15.0); // 5 * 3 * 1
        assert_eq!(double_factorial(6), 48.0); // 6 * 4 * 2
    }

    #[test]
    fn test_combination() {
        assert_eq!(combination(5, 0), 1.0);
        assert_eq!(combination(5, 5), 1.0);
        assert_eq!(combination(5, 2), 10.0);
        assert_eq!(combination(10, 3), 120.0);
        // C(n, k) = C(n, n-k)
        assert_eq!(combination(10, 3), combination(10, 7));
    }

    #[test]
    fn test_permutation() {
        assert_eq!(permutation(5, 0), 1.0);
        assert_eq!(permutation(5, 1), 5.0);
        assert_eq!(permutation(5, 2), 20.0);
        assert_eq!(permutation(10, 3), 720.0);
    }

    #[test]
    fn test_extend_binary() {
        assert_eq!(extend_binary("101", 5, '0'), "00101");
        assert_eq!(extend_binary("101", 3, '0'), "101");
        assert_eq!(extend_binary("101", 5, '1'), "11101");
    }

    #[test]
    fn test_is_negative_binary() {
        assert!(is_negative_binary("1000000000")); // 10 bits starting with 1
        assert!(!is_negative_binary("0111111111")); // 10 bits starting with 0
        assert!(!is_negative_binary("101")); // Less than 10 bits
    }

    #[test]
    fn test_twos_complement_value() {
        // The function calculates (2^bits + number) for two's complement
        // Positive 5 with 8 bits: 256 + 5 = 261
        assert_eq!(twos_complement_value(5, 8), 261);
        // Negative number: -1 with 8 bits: 256 + (-1) = 255
        assert_eq!(twos_complement_value(-1, 8), 255);
        // Negative number: -5 with 8 bits: 256 + (-5) = 251
        assert_eq!(twos_complement_value(-5, 8), 251);
    }

    #[test]
    fn test_parse_decimal_for_conversion_valid() {
        let value = CellValue::Int(100);
        let result = parse_decimal_for_conversion(&value, "TEST", -100, 100);
        assert_eq!(result.unwrap(), 100);
    }

    #[test]
    fn test_parse_decimal_for_conversion_out_of_range() {
        let value = CellValue::Int(200);
        let result = parse_decimal_for_conversion(&value, "TEST", -100, 100);
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_decimal_for_conversion_non_numeric() {
        let value = CellValue::String("abc".to_string());
        let result = parse_decimal_for_conversion(&value, "TEST", -100, 100);
        assert!(result.is_err());
    }

    #[test]
    fn test_ensure_number_in_range() {
        assert!(ensure_number_in_range(50, 0, 100, "TEST").is_ok());
        assert!(ensure_number_in_range(-1, 0, 100, "TEST").is_err());
        assert!(ensure_number_in_range(101, 0, 100, "TEST").is_err());
    }

    #[test]
    fn test_bit_operand_value_valid() {
        let value = CellValue::Int(42);
        let result = bit_operand_value(&value, "BITAND");
        assert_eq!(result.unwrap(), 42);
    }

    #[test]
    fn test_bit_operand_value_negative() {
        let value = CellValue::Int(-1);
        let result = bit_operand_value(&value, "BITAND");
        assert!(result.is_err());
    }

    #[test]
    fn test_bit_operand_value_non_numeric() {
        let value = CellValue::String("abc".to_string());
        let result = bit_operand_value(&value, "BITAND");
        assert!(result.is_err());
    }

    #[test]
    fn test_bit_shift_value_valid() {
        let value = CellValue::Int(10);
        let result = bit_shift_value(&value, "BITLSHIFT");
        assert_eq!(result.unwrap(), 10);
    }

    #[test]
    fn test_bit_shift_value_too_large() {
        let value = CellValue::Int(100);
        let result = bit_shift_value(&value, "BITLSHIFT");
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_places_argument_valid() {
        let value = CellValue::Int(5);
        let result = parse_places_argument(&value, "DEC2BIN");
        assert_eq!(result.unwrap(), 5);
    }

    #[test]
    fn test_parse_places_argument_out_of_range() {
        let value = CellValue::Int(15);
        let result = parse_places_argument(&value, "DEC2BIN");
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_places_argument_non_numeric() {
        let value = CellValue::String("abc".to_string());
        let result = parse_places_argument(&value, "DEC2BIN");
        assert!(result.is_err());
    }

    #[test]
    fn test_pad_with_places_success() {
        let mut result = "101".to_string();
        assert!(pad_with_places(&mut result, Some(5), "DEC2BIN").is_ok());
        assert_eq!(result, "00101");
    }

    #[test]
    fn test_pad_with_places_too_small() {
        let mut result = "1010".to_string();
        assert!(pad_with_places(&mut result, Some(2), "DEC2BIN").is_err());
    }

    #[test]
    fn test_binary_string_from_value_valid() {
        let value = CellValue::String("1010".to_string());
        let result = binary_string_from_value(&value, "BIN2DEC");
        assert_eq!(result.unwrap(), "1010");
    }

    #[test]
    fn test_binary_string_from_value_invalid_characters() {
        let value = CellValue::String("102".to_string());
        let result = binary_string_from_value(&value, "BIN2DEC");
        assert!(result.is_err());
    }

    #[test]
    fn test_binary_string_from_value_int() {
        let value = CellValue::Int(101);
        let result = binary_string_from_value(&value, "BIN2DEC");
        assert_eq!(result.unwrap(), "101");
    }

    #[test]
    fn test_parse_signed_binary_positive() {
        let result = parse_signed_binary("101", "BIN2DEC").unwrap();
        assert_eq!(result, 5);
    }

    #[test]
    fn test_parse_signed_binary_negative() {
        // 10-bit two's complement: 1111111111 = -1
        let result = parse_signed_binary("1111111111", "BIN2DEC").unwrap();
        assert_eq!(result, -1);
    }

    #[test]
    fn test_parse_signed_binary_out_of_range() {
        let result = parse_signed_binary("101010101010", "BIN2DEC");
        assert!(result.is_err());
    }

    #[test]
    fn test_hex_string_from_value_valid() {
        let value = CellValue::String("A1F".to_string());
        let result = parse_hex_string(&value, "HEX2DEC").unwrap();
        assert_eq!(result, "A1F");
    }

    #[test]
    fn test_hex_string_from_value_negative() {
        let value = CellValue::String("-A1F".to_string());
        let result = parse_hex_string(&value, "HEX2DEC").unwrap();
        assert_eq!(result, "-A1F");
    }

    #[test]
    fn test_hex_string_from_value_invalid() {
        let value = CellValue::String("GHI".to_string());
        let result = parse_hex_string(&value, "HEX2DEC");
        assert!(result.is_err());
    }

    #[test]
    fn test_signed_hex_to_decimal_positive() {
        let result = signed_hex_to_decimal("A", "HEX2DEC").unwrap();
        assert_eq!(result, 10);
    }

    #[test]
    fn test_signed_hex_to_decimal_negative() {
        let result = signed_hex_to_decimal("-A", "HEX2DEC").unwrap();
        assert_eq!(result, -10);
    }

    #[test]
    fn test_octal_string_from_value_valid() {
        let value = CellValue::String("755".to_string());
        let result = parse_octal_string(&value, "OCT2DEC").unwrap();
        assert_eq!(result, "755");
    }

    #[test]
    fn test_octal_string_from_value_invalid() {
        let value = CellValue::String("789".to_string());
        let result = parse_octal_string(&value, "OCT2DEC");
        assert!(result.is_err());
    }

    #[test]
    fn test_signed_octal_to_decimal_positive() {
        let result = signed_octal_to_decimal("755", "OCT2DEC").unwrap();
        assert_eq!(result, 0o755);
    }

    #[test]
    fn test_signed_octal_to_decimal_negative() {
        let result = signed_octal_to_decimal("-755", "OCT2DEC").unwrap();
        assert_eq!(result, -0o755);
    }

    #[test]
    fn test_negative_binary_to_hex() {
        let result = negative_binary_to_hex("1000000000");
        assert_eq!(result, "FFFFFFFE00");
    }

    #[test]
    fn test_negative_binary_to_oct() {
        let result = negative_binary_to_oct("1000000000");
        assert_eq!(result, "7777777000");
    }
}
