use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_bool, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_valuetotext(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "VALUETOTEXT expects 1 or 2 arguments (value, [format])".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let format = if args.len() == 2 {
        let f_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match crate::sheet::eval::engine::to_number(&f_val) {
            Some(n) => n.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };

    match format {
        0 => Ok(CellValue::String(to_text(&value))),
        1 => {
            // Strict format: includes quotes for strings
            match value {
                CellValue::String(s) => Ok(CellValue::String(format!("\"{}\"", s))),
                CellValue::Bool(b) => Ok(CellValue::String(
                    if b { "TRUE" } else { "FALSE" }.to_string(),
                )),
                CellValue::Empty => Ok(CellValue::String(String::new())),
                CellValue::Error(e) => Ok(CellValue::String(e)),
                _ => Ok(CellValue::String(to_text(&value))),
            }
        },
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_arraytotext(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "ARRAYTOTEXT expects 1 or 2 arguments (array, [format])".to_string(),
        ));
    }

    let format = if args.len() == 2 {
        let f_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match crate::sheet::eval::engine::to_number(&f_val) {
            Some(n) => n.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };

    let mut values = Vec::new();
    crate::sheet::eval::engine::for_each_value_in_expr(ctx, current_sheet, &args[0], |v| {
        values.push(v.clone());
        Ok(())
    })
    .await?;

    if format == 0 {
        let text_values: Vec<String> = values.iter().map(to_text).collect();
        Ok(CellValue::String(text_values.join(", ")))
    } else if format == 1 {
        // Array constant format: {"a", 1, TRUE; "b", 2, FALSE}
        // For now, we'll just join them with commas and semi-colons if it's a range,
        // but since for_each_value_in_expr flattens, we'll just do a simple list.
        let text_values: Vec<String> = values
            .iter()
            .map(|v| match v {
                CellValue::String(s) => format!("\"{}\"", s),
                CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
                _ => to_text(v),
            })
            .collect();
        Ok(CellValue::String(format!("{{{}}}", text_values.join(", "))))
    } else {
        Ok(CellValue::Error("#VALUE!".to_string()))
    }
}

pub(crate) async fn eval_textbefore(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 6 {
        return Ok(CellValue::Error(
            "TEXTBEFORE expects 2 to 6 arguments".to_string(),
        ));
    }
    let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?);
    let delimiter = to_text(&evaluate_expression(ctx, current_sheet, &args[1]).await?);
    if delimiter.is_empty() {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let instance_num = if args.len() >= 3 {
        crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[2]).await?,
        )
        .unwrap_or(1.0) as i32
    } else {
        1
    };

    let match_mode = if args.len() >= 4 {
        crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[3]).await?,
        )
        .unwrap_or(0.0) as i32
    } else {
        0
    };

    let _match_end = if args.len() >= 5 {
        crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[4]).await?,
        )
        .unwrap_or(0.0) as i32
    } else {
        0
    };

    let _if_not_found = if args.len() >= 6 {
        Some(to_text(
            &evaluate_expression(ctx, current_sheet, &args[5]).await?,
        ))
    } else {
        None
    };

    let haystack = if match_mode == 1 {
        text.to_lowercase()
    } else {
        text.clone()
    };
    let needle = if match_mode == 1 {
        delimiter.to_lowercase()
    } else {
        delimiter.clone()
    };

    let mut indices = Vec::new();
    let mut start = 0;
    while let Some(pos) = haystack[start..].find(&needle) {
        indices.push(start + pos);
        start += pos + needle.len();
    }

    if indices.is_empty() {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    let idx = if instance_num > 0 {
        if instance_num as usize > indices.len() {
            return Ok(CellValue::Error("#N/A".to_string()));
        }
        indices[instance_num as usize - 1]
    } else if instance_num < 0 {
        let abs_inst = instance_num.unsigned_abs() as usize;
        if abs_inst > indices.len() {
            return Ok(CellValue::Error("#N/A".to_string()));
        }
        indices[indices.len() - abs_inst]
    } else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };

    Ok(CellValue::String(text[..idx].to_string()))
}

pub(crate) async fn eval_textafter(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 6 {
        return Ok(CellValue::Error(
            "TEXTAFTER expects 2 to 6 arguments".to_string(),
        ));
    }
    let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?);
    let delimiter = to_text(&evaluate_expression(ctx, current_sheet, &args[1]).await?);
    if delimiter.is_empty() {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let instance_num = if args.len() >= 3 {
        crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[2]).await?,
        )
        .unwrap_or(1.0) as i32
    } else {
        1
    };

    let match_mode = if args.len() >= 4 {
        crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[3]).await?,
        )
        .unwrap_or(0.0) as i32
    } else {
        0
    };

    let haystack = if match_mode == 1 {
        text.to_lowercase()
    } else {
        text.clone()
    };
    let needle = if match_mode == 1 {
        delimiter.to_lowercase()
    } else {
        delimiter.clone()
    };

    let mut indices = Vec::new();
    let mut start = 0;
    while let Some(pos) = haystack[start..].find(&needle) {
        indices.push(start + pos);
        start += pos + needle.len();
    }

    if indices.is_empty() {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    let idx = if instance_num > 0 {
        if instance_num as usize > indices.len() {
            return Ok(CellValue::Error("#N/A".to_string()));
        }
        indices[instance_num as usize - 1]
    } else if instance_num < 0 {
        let abs_inst = instance_num.unsigned_abs() as usize;
        if abs_inst > indices.len() {
            return Ok(CellValue::Error("#N/A".to_string()));
        }
        indices[indices.len() - abs_inst]
    } else {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    };

    Ok(CellValue::String(text[idx + delimiter.len()..].to_string()))
}

pub(crate) async fn eval_textsplit(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 6 {
        return Ok(CellValue::Error(
            "TEXTSPLIT expects 2 to 6 arguments".to_string(),
        ));
    }

    let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?);

    // Delimiters can be a single value or an array of values
    let mut col_delimiters = Vec::new();
    crate::sheet::eval::engine::for_each_value_in_expr(ctx, current_sheet, &args[1], |v| {
        let s = to_text(v);
        if !s.is_empty() {
            col_delimiters.push(s);
        }
        Ok(())
    })
    .await?;

    let mut row_delimiters = Vec::new();
    if args.len() >= 3 {
        crate::sheet::eval::engine::for_each_value_in_expr(ctx, current_sheet, &args[2], |v| {
            let s = to_text(v);
            if !s.is_empty() {
                row_delimiters.push(s);
            }
            Ok(())
        })
        .await?;
    }

    let ignore_empty = if args.len() >= 4 {
        to_bool(&evaluate_expression(ctx, current_sheet, &args[3]).await?)
    } else {
        false
    };

    let match_mode = if args.len() >= 5 {
        match crate::sheet::eval::engine::to_number(
            &evaluate_expression(ctx, current_sheet, &args[4]).await?,
        ) {
            Some(n) => n.trunc() as i32,
            None => 0,
        }
    } else {
        0
    };

    // Very simplified implementation of TEXTSPLIT for now:
    // It should ideally return an array/dynamic spill, but we'll return a joined string
    // or just the first part for now until dynamic arrays are fully supported in eval.
    // Excel returns an array. If we can't return an array yet, we'll return #VALUE! or the first result.

    let mut all_delimiters = col_delimiters;
    all_delimiters.extend(row_delimiters);

    if all_delimiters.is_empty() {
        return Ok(CellValue::String(text));
    }

    let mut parts = vec![text];
    for delim in all_delimiters {
        let mut new_parts = Vec::new();
        for part in parts {
            if match_mode == 1 {
                // Case-insensitive split is harder with standard split
                // Simplified: just use standard split for now
                new_parts.extend(part.split(&delim).map(|s| s.to_string()));
            } else {
                new_parts.extend(part.split(&delim).map(|s| s.to_string()));
            }
        }
        parts = new_parts;
    }

    if ignore_empty {
        parts.retain(|s| !s.is_empty());
    }

    // Since we don't have full dynamic array support in the current eval context return,
    // we'll return the first part or a comma-separated list as a fallback.
    if parts.is_empty() {
        Ok(CellValue::Error("#N/A".to_string()))
    } else {
        Ok(CellValue::String(parts[0].clone()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::Expr;

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    #[tokio::test]
    async fn test_eval_valuetotext_default() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_valuetotext(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("hello".to_string()));
    }

    #[tokio::test]
    async fn test_eval_valuetotext_strict() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Format 1 = strict, adds quotes around strings
        let args = vec![str_expr("hello"), Expr::Literal(CellValue::Int(1))];
        let result = eval_valuetotext(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("\"hello\"".to_string()));
    }

    #[tokio::test]
    async fn test_eval_valuetotext_bool() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Bool(true))];
        let result = eval_valuetotext(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("TRUE".to_string()));
    }

    #[tokio::test]
    async fn test_eval_valuetotext_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_valuetotext(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 or 2")),
            _ => panic!("Expected Error, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_arraytotext_default() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Set up a range with values
        for i in 0..3 {
            engine.set_cell("Sheet1", 0, i, CellValue::Int(i as i64 + 1));
        }
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 2,
        });
        let args = vec![range];
        let result = eval_arraytotext(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("1, 2, 3".to_string()));
    }

    #[tokio::test]
    async fn test_eval_arraytotext_array_format() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        engine.set_cell("Sheet1", 0, 0, CellValue::String("a".to_string()));
        engine.set_cell("Sheet1", 0, 1, CellValue::Int(1));
        let range = Expr::Range(crate::sheet::eval::parser::RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 1,
        });
        let args = vec![range, Expr::Literal(CellValue::Int(1))];
        let result = eval_arraytotext(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("{\"a\", 1}".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textbefore() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello-world-test"), str_expr("-")];
        let result = eval_textbefore(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("hello".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textbefore_instance() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Get text before 2nd hyphen
        let args = vec![
            str_expr("hello-world-test"),
            str_expr("-"),
            Expr::Literal(CellValue::Int(2)),
        ];
        let result = eval_textbefore(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("hello-world".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textbefore_not_found() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello"), str_expr("-")];
        let result = eval_textbefore(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#N/A".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textafter() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello-world-test"), str_expr("-")];
        let result = eval_textafter(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("world-test".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textafter_instance() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        // Get text after 2nd hyphen
        let args = vec![
            str_expr("hello-world-test"),
            str_expr("-"),
            Expr::Literal(CellValue::Int(2)),
        ];
        let result = eval_textafter(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("test".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textsplit() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("a,b,c"), str_expr(",")];
        let result = eval_textsplit(ctx, "Sheet1", &args).await.unwrap();
        // Returns first part since no dynamic array support
        assert_eq!(result, CellValue::String("a".to_string()));
    }

    #[tokio::test]
    async fn test_eval_textsplit_empty_delim() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("abc"), str_expr("")];
        let result = eval_textsplit(ctx, "Sheet1", &args).await.unwrap();
        // Empty delimiter returns original text
        assert_eq!(result, CellValue::String("abc".to_string()));
    }
}
