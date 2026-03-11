use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    char_index_from_dbcs_byte, dbcs_byte_prefixes, replace_bytes_segment, replace_chars_segment,
    take_left, take_left_bytes, take_mid, take_mid_bytes, take_right, take_right_bytes,
    to_non_negative_int, to_positive_int,
};

pub(crate) async fn eval_char(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CHAR expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let code = match to_positive_int(&value) {
        Some(n) if n <= 255 => n as u8,
        _ => {
            return Ok(CellValue::Error(
                "CHAR code must be an integer between 1 and 255".to_string(),
            ));
        },
    };
    Ok(CellValue::String((code as char).to_string()))
}

pub(crate) async fn eval_code(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CODE expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&value);
    if let Some(ch) = text.chars().next() {
        Ok(CellValue::Int(ch as u32 as i64))
    } else {
        Ok(CellValue::Error("CODE text must not be empty".to_string()))
    }
}

pub(crate) async fn eval_clean(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("CLEAN expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&value);
    let cleaned: String = text
        .chars()
        .filter(|c| {
            let code = *c as u32;
            code >= 32
        })
        .collect();
    Ok(CellValue::String(cleaned))
}

pub(crate) async fn eval_left(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "LEFT expects 1 or 2 arguments (text, [num_chars])".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let num_chars = if args.len() == 2 {
        let num_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_non_negative_int(&num_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "LEFT num_chars must be a non-negative integer".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let result = take_left(&s, num_chars);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_leftb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "LEFTB expects 1 or 2 arguments (text, [num_bytes])".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let num_bytes = if args.len() == 2 {
        let num_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_non_negative_int(&num_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "LEFTB num_bytes must be a non-negative integer".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let result = take_left_bytes(&s, num_bytes);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_right(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "RIGHT expects 1 or 2 arguments (text, [num_chars])".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let num_chars = if args.len() == 2 {
        let num_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_non_negative_int(&num_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "RIGHT num_chars must be a non-negative integer".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let result = take_right(&s, num_chars);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_rightb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "RIGHTB expects 1 or 2 arguments (text, [num_bytes])".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let num_bytes = if args.len() == 2 {
        let num_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_non_negative_int(&num_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "RIGHTB num_bytes must be a non-negative integer".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let result = take_right_bytes(&s, num_bytes);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_mid(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "MID expects 3 arguments (text, start_num, num_chars)".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let start_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let start_num = match to_positive_int(&start_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "MID start_num must be a positive integer".to_string(),
            ));
        },
    };
    let count_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let num_chars = match to_non_negative_int(&count_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "MID num_chars must be a non-negative integer".to_string(),
            ));
        },
    };
    let result = take_mid(&s, start_num, num_chars);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_midb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "MIDB expects 3 arguments (text, start_num, num_bytes)".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let s = to_text(&text);
    let start_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let start_num = match to_positive_int(&start_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "MIDB start_num must be a positive integer".to_string(),
            ));
        },
    };
    let count_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let num_bytes = match to_non_negative_int(&count_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "MIDB num_bytes must be a non-negative integer".to_string(),
            ));
        },
    };
    let result = take_mid_bytes(&s, start_num, num_bytes);
    Ok(CellValue::String(result))
}

pub(crate) async fn eval_replace(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "REPLACE expects 4 arguments (old_text, start_num, num_chars, new_text)".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let original = to_text(&text);
    let start_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let start_num = match to_positive_int(&start_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "REPLACE start_num must be a positive integer".to_string(),
            ));
        },
    };
    let count_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let num_chars = match to_non_negative_int(&count_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "REPLACE num_chars must be a non-negative integer".to_string(),
            ));
        },
    };
    let new_text = evaluate_expression(ctx, current_sheet, &args[3]).await?;
    let replacement = to_text(&new_text);

    if start_num > original.chars().count() {
        return Ok(CellValue::String(original));
    }

    let replaced = replace_chars_segment(&original, start_num, num_chars, &replacement)
        .unwrap_or_else(|| original.clone());
    Ok(CellValue::String(replaced))
}

pub(crate) async fn eval_replaceb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 4 {
        return Ok(CellValue::Error(
            "REPLACEB expects 4 arguments (old_text, start_num, num_bytes, new_text)".to_string(),
        ));
    }
    let text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let original = to_text(&text);
    let start_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let start_byte = match to_positive_int(&start_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "REPLACEB start_num must be a positive integer".to_string(),
            ));
        },
    };
    let count_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
    let num_bytes = match to_non_negative_int(&count_val) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "REPLACEB num_bytes must be a non-negative integer".to_string(),
            ));
        },
    };
    let new_text = evaluate_expression(ctx, current_sheet, &args[3]).await?;
    let replacement = to_text(&new_text);

    let total_bytes = super::helpers::dbcs_byte_len(&original);
    if start_byte > total_bytes {
        return Ok(CellValue::String(original));
    }

    let replaced = replace_bytes_segment(&original, start_byte, num_bytes, &replacement)
        .unwrap_or_else(|| original.clone());
    Ok(CellValue::String(replaced))
}

pub(crate) async fn eval_find(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_find_like(ctx, current_sheet, args, true, "FIND").await
}

pub(crate) async fn eval_findb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "FINDB expects 2 or 3 arguments (find_text, within_text, [start_num])".to_string(),
        ));
    }
    let find_text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let needle = to_text(&find_text);
    let within_text = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let haystack = to_text(&within_text);
    let start_byte = if args.len() == 3 {
        let start_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_positive_int(&start_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "FINDB start_num must be a positive integer".to_string(),
                ));
            },
        }
    } else {
        1
    };

    let prefixes = dbcs_byte_prefixes(&haystack);
    let total_bytes = *prefixes.last().unwrap_or(&0);
    if start_byte == 0 || start_byte > total_bytes.saturating_add(1) {
        return Ok(CellValue::Error(
            "FINDB start_num is out of range".to_string(),
        ));
    }
    if needle.is_empty() {
        return Ok(CellValue::Int(start_byte as i64));
    }

    let hay_chars: Vec<(usize, char)> = haystack.char_indices().collect();
    let hay_len = hay_chars.len();
    let start_idx = if start_byte > total_bytes {
        hay_len
    } else {
        char_index_from_dbcs_byte(&prefixes, start_byte).unwrap_or(hay_len)
    };

    let needle_char_len = needle.chars().count();
    if hay_len < needle_char_len || start_idx >= hay_len {
        return Ok(CellValue::Error(
            "FINDB could not find the text".to_string(),
        ));
    }

    for i in start_idx..=hay_len - needle_char_len {
        let start_byte_offset = hay_chars[i].0;
        let end_byte_offset = if i + needle_char_len < hay_len {
            hay_chars[i + needle_char_len].0
        } else {
            haystack.len()
        };
        let candidate = &haystack[start_byte_offset..end_byte_offset];
        if candidate == needle {
            let byte_position = prefixes[i] + 1;
            return Ok(CellValue::Int(byte_position as i64));
        }
    }

    Ok(CellValue::Error(
        "FINDB could not find the text".to_string(),
    ))
}

pub(crate) async fn eval_search(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_find_like(ctx, current_sheet, args, false, "SEARCH").await
}

pub(crate) async fn eval_searchb(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "SEARCHB expects 2 or 3 arguments (find_text, within_text, [start_num])".to_string(),
        ));
    }
    let find_text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let needle = to_text(&find_text);
    let within_text = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let haystack = to_text(&within_text);
    let start_byte = if args.len() == 3 {
        let start_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_positive_int(&start_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(
                    "SEARCHB start_num must be a positive integer".to_string(),
                ));
            },
        }
    } else {
        1
    };

    let prefixes = dbcs_byte_prefixes(&haystack);
    let total_bytes = *prefixes.last().unwrap_or(&0);
    if start_byte == 0 || start_byte > total_bytes.saturating_add(1) {
        return Ok(CellValue::Error(
            "SEARCHB start_num is out of range".to_string(),
        ));
    }
    if needle.is_empty() {
        return Ok(CellValue::Int(start_byte as i64));
    }

    let hay_chars: Vec<(usize, char)> = haystack.char_indices().collect();
    let hay_len = hay_chars.len();
    let start_idx = if start_byte > total_bytes {
        hay_len
    } else {
        char_index_from_dbcs_byte(&prefixes, start_byte).unwrap_or(hay_len)
    };

    let needle_char_len = needle.chars().count();
    if hay_len < needle_char_len || start_idx >= hay_len {
        return Ok(CellValue::Error(
            "SEARCHB could not find the text".to_string(),
        ));
    }

    let needle_cmp = needle.to_lowercase();
    for i in start_idx..=hay_len - needle_char_len {
        let start_byte_offset = hay_chars[i].0;
        let end_byte_offset = if i + needle_char_len < hay_len {
            hay_chars[i + needle_char_len].0
        } else {
            haystack.len()
        };
        let candidate = &haystack[start_byte_offset..end_byte_offset];
        if candidate.to_lowercase() == needle_cmp {
            let byte_position = prefixes[i] + 1;
            return Ok(CellValue::Int(byte_position as i64));
        }
    }

    Ok(CellValue::Error(
        "SEARCHB could not find the text".to_string(),
    ))
}

async fn eval_find_like(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    case_sensitive: bool,
    name: &str,
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(format!(
            "{name} expects 2 or 3 arguments (find_text, within_text, [start_num])"
        )));
    }
    let find_text = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let needle = to_text(&find_text);
    if needle.is_empty() {
        return Ok(CellValue::Error(format!(
            "{name} requires non-empty find_text"
        )));
    }
    let within_text = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let haystack = to_text(&within_text);
    let start_num = if args.len() == 3 {
        let start_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_positive_int(&start_val) {
            Some(n) => n,
            None => {
                return Ok(CellValue::Error(format!(
                    "{name} start_num must be a positive integer"
                )));
            },
        }
    } else {
        1
    };
    let hay_chars: Vec<(usize, char)> = haystack.char_indices().collect();
    let hay_len = hay_chars.len();
    if start_num == 0 || start_num > hay_len.saturating_add(1) {
        return Ok(CellValue::Error(format!(
            "{name} start_num is out of range"
        )));
    }
    let start_idx = start_num.saturating_sub(1);
    let needle_char_len = needle.chars().count();
    if needle_char_len == 0 {
        return Ok(CellValue::Int(start_num as i64));
    }
    if hay_len < needle_char_len || start_idx >= hay_len {
        return Ok(CellValue::Error(format!("{name} could not find the text")));
    }
    for i in start_idx..=hay_len - needle_char_len {
        let start_byte = hay_chars[i].0;
        let end_byte = if i + needle_char_len < hay_len {
            hay_chars[i + needle_char_len].0
        } else {
            haystack.len()
        };
        let candidate = &haystack[start_byte..end_byte];
        let matched = if case_sensitive {
            candidate == needle
        } else {
            candidate.to_lowercase() == needle.to_lowercase()
        };
        if matched {
            return Ok(CellValue::Int((i + 1) as i64));
        }
    }
    Ok(CellValue::Error(format!("{name} could not find the text")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    #[tokio::test]
    async fn test_eval_char() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(65.0)];
        let result = eval_char(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "A"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_char_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(300.0)];
        let result = eval_char(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("between 1 and 255")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_code() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("A")];
        let result = eval_code(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 65),
            _ => panic!("Expected Int(65)"),
        }
    }

    #[tokio::test]
    async fn test_eval_code_empty() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("")];
        let result = eval_code(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("must not be empty")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_clean() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello\x00world")];
        let result = eval_clean(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "helloworld"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_left() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello"), num_expr(3.0)];
        let result = eval_left(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "hel"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_left_default() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_left(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "h"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_right() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello"), num_expr(3.0)];
        let result = eval_right(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "llo"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_mid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello world"), num_expr(7.0), num_expr(5.0)];
        let result = eval_mid(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "world"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_mid_start_at_1() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello"), num_expr(1.0), num_expr(3.0)];
        let result = eval_mid(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "hel"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_replace() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            str_expr("hello world"),
            num_expr(7.0),
            num_expr(5.0),
            str_expr("universe"),
        ];
        let result = eval_replace(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "hello universe"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_find() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("world"), str_expr("hello world")];
        let result = eval_find(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 7),
            _ => panic!("Expected Int(7)"),
        }
    }

    #[tokio::test]
    async fn test_eval_find_case_sensitive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("WORLD"), str_expr("hello world")];
        let result = eval_find(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("could not find")),
            _ => panic!("Expected Error - FIND is case sensitive"),
        }
    }

    #[tokio::test]
    async fn test_eval_search() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("world"), str_expr("hello world")];
        let result = eval_search(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 7),
            _ => panic!("Expected Int(7)"),
        }
    }

    #[tokio::test]
    async fn test_eval_search_case_insensitive() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("WORLD"), str_expr("hello world")];
        let result = eval_search(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(v) => assert_eq!(v, 7),
            _ => panic!("Expected Int(7) - SEARCH is case insensitive"),
        }
    }
}
