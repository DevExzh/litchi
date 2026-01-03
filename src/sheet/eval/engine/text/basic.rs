use crate::sheet::eval::engine::{
    EvalCtx, evaluate_expression, for_each_value_in_expr, is_blank, to_bool, to_text,
};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{to_non_negative_int, to_positive_int};

pub(crate) async fn eval_len(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("LEN expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v);
    Ok(CellValue::Int(s.chars().count() as i64))
}

fn fullwidth_to_halfwidth(s: &str) -> String {
    s.chars()
        .map(|ch| match ch {
            '\u{FF01}'..='\u{FF5E}' => char::from_u32((ch as u32) - 0xFEE0).unwrap_or(ch),
            '\u{3000}' => ' ',
            _ => ch,
        })
        .collect()
}

fn halfwidth_to_fullwidth(s: &str) -> String {
    s.chars()
        .map(|ch| match ch {
            '\u{0021}'..='\u{007E}' => char::from_u32((ch as u32) + 0xFEE0).unwrap_or(ch),
            ' ' => '\u{3000}',
            _ => ch,
        })
        .collect()
}

pub(crate) async fn eval_asc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ASC expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&value);
    Ok(CellValue::String(fullwidth_to_halfwidth(&text)))
}

pub(crate) async fn eval_jis(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("JIS expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&value);
    Ok(CellValue::String(halfwidth_to_fullwidth(&text)))
}

pub(crate) async fn eval_proper(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("PROPER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v);
    let mut result = String::with_capacity(s.len());
    let mut new_word = true;
    for ch in s.chars() {
        if ch.is_alphabetic() {
            if new_word {
                for upper in ch.to_uppercase() {
                    result.push(upper);
                }
            } else {
                for lower in ch.to_lowercase() {
                    result.push(lower);
                }
            }
            new_word = false;
        } else {
            new_word = true;
            result.push(ch);
        }
    }
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_lenb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("LENB expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v);
    Ok(CellValue::Int(s.len() as i64))
}

pub(crate) async fn eval_lower(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("LOWER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v).to_lowercase();
    Ok(CellValue::String(s))
}

pub(crate) async fn eval_upper(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UPPER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v).to_uppercase();
    Ok(CellValue::String(s))
}

pub(crate) async fn eval_trim(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("TRIM expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&v);
    Ok(CellValue::String(s.trim().to_string()))
}

pub(crate) async fn eval_concat(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut out = String::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            out.push_str(&to_text(v));
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::String(out))
}

pub(crate) async fn eval_textjoin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 {
        return Ok(CellValue::Error(
            "TEXTJOIN expects at least 3 arguments (delimiter, ignore_empty, text1, ...)"
                .to_string(),
        ));
    }

    let delimiter_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let delimiter = to_text(&delimiter_val);

    let ignore_empty_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let ignore_empty = to_bool(&ignore_empty_val);

    let mut parts: Vec<String> = Vec::new();
    for arg in &args[2..] {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if ignore_empty && is_blank(v) {
                return Ok(());
            }
            parts.push(to_text(v));
            Ok(())
        })
        .await?;
    }

    Ok(CellValue::String(parts.join(&delimiter)))
}

pub(crate) async fn eval_rept(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "REPT expects 2 arguments (text, number_times)".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let times_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let times = match to_non_negative_int(&times_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "REPT number_times must be a non-negative integer".to_string(),
            ));
        },
    };
    const MAX_REPT_RESULT: usize = 32_767;
    if s.is_empty() || times == 0 {
        return Ok(CellValue::String(String::new()));
    }
    if s.chars().count().saturating_mul(times) > MAX_REPT_RESULT {
        return Ok(CellValue::Error(
            "REPT result exceeds maximum length of 32767 characters".to_string(),
        ));
    }
    let mut out = String::with_capacity(s.len().saturating_mul(times));
    for _ in 0..times {
        out.push_str(&s);
    }
    Ok(CellValue::String(out))
}

pub(crate) async fn eval_exact(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EXACT expects 2 arguments (text1, text2)".to_string(),
        ));
    }
    let left = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let right = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    Ok(CellValue::Bool(to_text(&left) == to_text(&right)))
}

pub(crate) async fn eval_phonetic(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("PHONETIC expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match value {
        CellValue::String(_) => Ok(value),
        CellValue::Error(_) => Ok(value),
        _ => Ok(CellValue::String(String::new())),
    }
}

fn numbervalue_to_cell(value: f64) -> CellValue {
    if value.fract().abs() < f64::EPSILON && value.abs() <= i64::MAX as f64 {
        CellValue::Int(value as i64)
    } else {
        CellValue::Float(value)
    }
}

fn parse_numbervalue(
    text: &str,
    decimal: char,
    group: Option<char>,
) -> std::result::Result<f64, String> {
    if let Some(g) = group
        && g == decimal
    {
        return Err("#VALUE!".to_string());
    }
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Err("#VALUE!".to_string());
    }
    let mut cleaned = String::with_capacity(trimmed.len());
    let mut chars = trimmed.chars().peekable();
    if let Some(sign) = chars.peek().copied()
        && (sign == '+' || sign == '-')
    {
        cleaned.push(sign);
        chars.next();
    }
    let mut decimal_seen = false;
    for ch in chars {
        if ch == decimal {
            if decimal_seen {
                return Err("#VALUE!".to_string());
            }
            cleaned.push('.');
            decimal_seen = true;
            continue;
        }
        if let Some(g) = group
            && ch == g
        {
            continue;
        }
        if ch.is_ascii_digit() {
            cleaned.push(ch);
        } else {
            return Err("#VALUE!".to_string());
        }
    }
    if cleaned.is_empty()
        || cleaned == "+"
        || cleaned == "-"
        || cleaned == "+."
        || cleaned == "-."
        || cleaned == "."
    {
        return Err("#VALUE!".to_string());
    }
    cleaned.parse::<f64>().map_err(|_| "#VALUE!".to_string())
}

pub(crate) async fn eval_numbervalue(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 3 {
        return Ok(CellValue::Error(
            "NUMBERVALUE expects 1 to 3 arguments (text, [decimal_separator], [group_separator])"
                .to_string(),
        ));
    }
    let text_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&text_val);

    let decimal_char = if args.len() >= 2 {
        let dec_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        let dec_text = to_text(&dec_val);
        let trimmed = dec_text.trim();
        if trimmed.len() != 1 {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        trimmed.chars().next().unwrap()
    } else {
        '.'
    };

    let group_char = if args.len() == 3 {
        let group_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        let group_text = to_text(&group_val);
        let trimmed = group_text.trim();
        if trimmed.is_empty() {
            None
        } else if trimmed.len() == 1 {
            Some(trimmed.chars().next().unwrap())
        } else {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
    } else {
        None
    };

    match parse_numbervalue(&text, decimal_char, group_char) {
        Ok(value) => Ok(numbervalue_to_cell(value)),
        Err(err) => Ok(CellValue::Error(err)),
    }
}

pub(crate) async fn eval_substitute(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 || args.len() > 4 {
        return Ok(CellValue::Error(
            "SUBSTITUTE expects 3 or 4 arguments (text, old_text, new_text, [instance_num])"
                .to_string(),
        ));
    }
    let text_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&text_val);
    let old_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let old_text = to_text(&old_val);
    if old_text.is_empty() {
        return Ok(CellValue::Error(
            "SUBSTITUTE old_text must not be empty".to_string(),
        ));
    }
    let new_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let new_text = to_text(&new_val);
    if args.len() == 4 {
        let instance_val = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        let instance = match to_positive_int(&instance_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "SUBSTITUTE instance_num must be a positive integer".to_string(),
                ));
            },
        };
        let mut occurrence = 0;
        let mut index = 0;
        while let Some(pos) = text[index..].find(&old_text) {
            let absolute = index + pos;
            occurrence += 1;
            if occurrence == instance {
                let mut result = String::with_capacity(
                    text.len() + new_text.len().saturating_sub(old_text.len()),
                );
                result.push_str(&text[..absolute]);
                result.push_str(&new_text);
                result.push_str(&text[absolute + old_text.len()..]);
                return Ok(CellValue::String(result));
            }
            index = absolute + old_text.len();
        }
        Ok(CellValue::String(text))
    } else {
        Ok(CellValue::String(text.replace(&old_text, &new_text)))
    }
}
