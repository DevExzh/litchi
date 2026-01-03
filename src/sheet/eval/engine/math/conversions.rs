use crate::sheet::{CellValue, Result};

use super::helpers::{
    BIN_MAX, BIN_MIN, HEX_MAX, HEX_MIN, OCT_MAX, OCT_MIN, binary_string_from_value,
    ensure_number_in_range, is_negative_binary, negative_binary_to_hex, negative_binary_to_oct,
    pad_with_places, parse_decimal_for_conversion, parse_hex_string, parse_octal_string,
    parse_places_argument, parse_signed_binary, signed_hex_to_decimal, signed_octal_to_decimal,
    to_int_if_whole, twos_complement_value,
};
use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number, to_text};
use crate::sheet::eval::parser::Expr;

pub(crate) async fn eval_bin2dec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("BIN2DEC expects 1 argument".to_string()));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let bits = match binary_string_from_value(&number_value, "BIN2DEC") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let value = match parse_signed_binary(&bits, "BIN2DEC") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int(value))
}

pub(crate) async fn eval_bin2hex(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "BIN2HEX expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "BIN2HEX") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    let bits = match binary_string_from_value(&number_value, "BIN2HEX") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let result = if is_negative_binary(&bits) {
        negative_binary_to_hex(&bits)
    } else {
        let value = u32::from_str_radix(&bits, 2).unwrap();
        let mut hex = format!("{value:X}");
        if let Err(err) = pad_with_places(&mut hex, places, "BIN2HEX") {
            return Ok(err);
        }
        hex
    };
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_bin2oct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "BIN2OCT expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "BIN2OCT") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    let bits = match binary_string_from_value(&number_value, "BIN2OCT") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let result = if is_negative_binary(&bits) {
        negative_binary_to_oct(&bits)
    } else {
        let value = u32::from_str_radix(&bits, 2).unwrap();
        let mut oct = format!("{value:o}");
        if let Err(err) = pad_with_places(&mut oct, places, "BIN2OCT") {
            return Ok(err);
        }
        oct
    };
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_dec2bin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "DEC2BIN expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let number = match parse_decimal_for_conversion(&number_value, "DEC2BIN", BIN_MIN, BIN_MAX) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "DEC2BIN") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if number < 0 {
        let value = twos_complement_value(number, 10);
        return Ok(CellValue::String(format!("{value:010b}")));
    }
    let mut result = format!("{number:b}");
    if let Err(err) = pad_with_places(&mut result, places, "DEC2BIN") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

const DECIMAL_MAX_VALUE: u128 = (1u128 << 53) - 1;

pub(crate) async fn eval_decimal(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "DECIMAL expects 2 arguments (text, radix)".to_string(),
        ));
    }

    let text_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let raw_text = to_text(&text_val);
    let text = raw_text.trim();
    if text.is_empty() {
        return Ok(CellValue::Error(
            "DECIMAL text must not be empty".to_string(),
        ));
    }
    if text.len() > 255 {
        return Ok(CellValue::Error(
            "DECIMAL text must be 255 characters or fewer".to_string(),
        ));
    }

    let radix_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let radix_num = match to_number(&radix_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "DECIMAL radix must be numeric".to_string(),
            ));
        },
    };
    let radix_int = match to_int_if_whole(radix_num) {
        Some(i) => i,
        None => {
            return Ok(CellValue::Error(
                "DECIMAL radix must be an integer between 2 and 36".to_string(),
            ));
        },
    };
    if !(2..=36).contains(&radix_int) {
        return Ok(CellValue::Error(
            "DECIMAL radix must be between 2 and 36".to_string(),
        ));
    }
    let radix = radix_int as u32;

    let mut value: u128 = 0;
    for ch in text.chars() {
        let digit = match char_to_digit(ch) {
            Some(d) => d,
            None => {
                return Ok(CellValue::Error(format!(
                    "DECIMAL text contains invalid character '{}'",
                    ch
                )));
            },
        };
        if digit >= radix {
            return Ok(CellValue::Error(format!(
                "DECIMAL text contains digit '{}' invalid for radix {}",
                ch, radix
            )));
        }
        value = value * radix as u128 + digit as u128;
        if value > DECIMAL_MAX_VALUE {
            return Ok(CellValue::Error(
                "DECIMAL result is out of supported range (must be less than 2^53)".to_string(),
            ));
        }
    }

    Ok(CellValue::Int(value as i64))
}

pub(crate) async fn eval_base(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "BASE expects 2 or 3 arguments (number, radix, [min_length])".to_string(),
        ));
    }

    let number_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number_num = match to_number(&number_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("BASE number must be numeric".to_string()));
        },
    };
    let number_int = match to_int_if_whole(number_num) {
        Some(i) => i,
        None => {
            return Ok(CellValue::Error(
                "BASE number must be an integer between 0 and 2^53-1".to_string(),
            ));
        },
    };
    if number_int < 0 || number_int as u128 > DECIMAL_MAX_VALUE {
        return Ok(CellValue::Error(
            "BASE number must be between 0 and 2^53-1".to_string(),
        ));
    }
    let number = number_int as u128;

    let radix_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let radix_num = match to_number(&radix_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("BASE radix must be numeric".to_string()));
        },
    };
    let radix_int = match to_int_if_whole(radix_num) {
        Some(i) => i,
        None => {
            return Ok(CellValue::Error(
                "BASE radix must be an integer between 2 and 36".to_string(),
            ));
        },
    };
    if !(2..=36).contains(&radix_int) {
        return Ok(CellValue::Error(
            "BASE radix must be between 2 and 36".to_string(),
        ));
    }
    let radix = radix_int as u32;

    let min_length = if args.len() == 3 {
        let len_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        let len_num = match to_number(&len_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "BASE min_length must be numeric".to_string(),
                ));
            },
        };
        let len_int = match to_int_if_whole(len_num) {
            Some(i) => i,
            None => {
                return Ok(CellValue::Error(
                    "BASE min_length must be a non-negative integer up to 255".to_string(),
                ));
            },
        };
        if !(0..=255).contains(&len_int) {
            return Ok(CellValue::Error(
                "BASE min_length must be between 0 and 255".to_string(),
            ));
        }
        len_int as usize
    } else {
        0
    };

    let mut result = if number == 0 {
        "0".to_string()
    } else {
        convert_number_to_base(number, radix)
    };

    if min_length > result.len() {
        let zeros = "0".repeat(min_length - result.len());
        result = format!("{zeros}{result}");
    }

    Ok(CellValue::String(result))
}

fn char_to_digit(c: char) -> Option<u32> {
    if c.is_ascii_digit() {
        Some(c as u32 - '0' as u32)
    } else if c.is_ascii_alphabetic() {
        let upper = c.to_ascii_uppercase();
        Some(upper as u32 - 'A' as u32 + 10)
    } else {
        None
    }
}

fn convert_number_to_base(mut value: u128, radix: u32) -> String {
    debug_assert!((2..=36).contains(&radix));
    let mut digits: Vec<char> = Vec::new();
    while value > 0 {
        let rem = (value % radix as u128) as u32;
        digits.push(digit_to_char(rem));
        value /= radix as u128;
    }
    digits.iter().rev().collect()
}

fn digit_to_char(value: u32) -> char {
    if value < 10 {
        (b'0' + value as u8) as char
    } else {
        (b'A' + (value - 10) as u8) as char
    }
}

pub(crate) async fn eval_dec2oct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "DEC2OCT expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let number = match parse_decimal_for_conversion(&number_value, "DEC2OCT", OCT_MIN, OCT_MAX) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "DEC2OCT") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if number < 0 {
        let value = twos_complement_value(number, 30);
        return Ok(CellValue::String(format!("{value:010o}")));
    }
    let mut result = format!("{number:o}");
    if let Err(err) = pad_with_places(&mut result, places, "DEC2OCT") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_dec2hex(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "DEC2HEX expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let number = match parse_decimal_for_conversion(&number_value, "DEC2HEX", HEX_MIN, HEX_MAX) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "DEC2HEX") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if number < 0 {
        let value = twos_complement_value(number, 40);
        return Ok(CellValue::String(format!("{value:010X}")));
    }
    let mut result = format!("{number:X}");
    if let Err(err) = pad_with_places(&mut result, places, "DEC2HEX") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_hex2dec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("HEX2DEC expects 1 argument".to_string()));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let hex = match parse_hex_string(&number_value, "HEX2DEC") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let value = match signed_hex_to_decimal(&hex, "HEX2DEC") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int(value))
}

pub(crate) async fn eval_hex2bin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "HEX2BIN expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let hex = match parse_hex_string(&number_value, "HEX2BIN") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let decimal = match signed_hex_to_decimal(&hex, "HEX2BIN") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    if let Err(err) = ensure_number_in_range(decimal, BIN_MIN, BIN_MAX, "HEX2BIN") {
        return Ok(err);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "HEX2BIN") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if decimal < 0 {
        let value = twos_complement_value(decimal, 10);
        return Ok(CellValue::String(format!("{value:010b}")));
    }
    let mut result = format!("{decimal:b}");
    if let Err(err) = pad_with_places(&mut result, places, "HEX2BIN") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_hex2oct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "HEX2OCT expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let hex = match parse_hex_string(&number_value, "HEX2OCT") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let decimal = match signed_hex_to_decimal(&hex, "HEX2OCT") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    if let Err(err) = ensure_number_in_range(decimal, OCT_MIN, OCT_MAX, "HEX2OCT") {
        return Ok(err);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "HEX2OCT") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if decimal < 0 {
        let value = twos_complement_value(decimal, 30);
        return Ok(CellValue::String(format!("{value:010o}")));
    }
    let mut result = format!("{decimal:o}");
    if let Err(err) = pad_with_places(&mut result, places, "HEX2OCT") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_oct2dec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("OCT2DEC expects 1 argument".to_string()));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let oct = match parse_octal_string(&number_value, "OCT2DEC") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let value = match signed_octal_to_decimal(&oct, "OCT2DEC") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Int(value))
}

pub(crate) async fn eval_oct2bin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "OCT2BIN expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let oct = match parse_octal_string(&number_value, "OCT2BIN") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let decimal = match signed_octal_to_decimal(&oct, "OCT2BIN") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    if let Err(err) = ensure_number_in_range(decimal, BIN_MIN, BIN_MAX, "OCT2BIN") {
        return Ok(err);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "OCT2BIN") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if decimal < 0 {
        let value = twos_complement_value(decimal, 10);
        return Ok(CellValue::String(format!("{value:010b}")));
    }
    let mut result = format!("{decimal:b}");
    if let Err(err) = pad_with_places(&mut result, places, "OCT2BIN") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_oct2hex(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "OCT2HEX expects 1 or 2 arguments".to_string(),
        ));
    }
    let number_value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = number_value {
        return Ok(number_value);
    }
    let oct = match parse_octal_string(&number_value, "OCT2HEX") {
        Ok(s) => s,
        Err(err) => return Ok(err),
    };
    let decimal = match signed_octal_to_decimal(&oct, "OCT2HEX") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    if let Err(err) = ensure_number_in_range(decimal, HEX_MIN, HEX_MAX, "OCT2HEX") {
        return Ok(err);
    }
    let places = if args.len() == 2 {
        let places_value = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        if let CellValue::Error(_) = places_value {
            return Ok(places_value);
        }
        Some(match parse_places_argument(&places_value, "OCT2HEX") {
            Ok(p) => p,
            Err(err) => return Ok(err),
        })
    } else {
        None
    };
    if decimal < 0 {
        let value = twos_complement_value(decimal, 40);
        return Ok(CellValue::String(format!("{value:010X}")));
    }
    let mut result = format!("{decimal:X}");
    if let Err(err) = pad_with_places(&mut result, places, "OCT2HEX") {
        return Ok(err);
    }
    Ok(CellValue::String(result))
}
