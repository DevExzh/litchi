use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EvalCtx, ResolvedName, evaluate_expression, is_blank, to_number};

const EPS: f64 = 1e-12;

pub(crate) fn error_code(value: &CellValue) -> Option<&str> {
    match value {
        CellValue::Error(code) => Some(code.as_str()),
        _ => None,
    }
}

pub(crate) fn is_na_error(code: &str) -> bool {
    code.eq_ignore_ascii_case("#N/A")
}

fn is_even_number(value: f64) -> bool {
    ((value / 2.0).fract()).abs() < EPS
}

pub(crate) enum ReferenceKind {
    Single { sheet: String, row: u32, col: u32 },
    Range,
    None,
    Error(CellValue),
}

fn single_cell_from_range(
    sheet: &str,
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
) -> Option<(String, u32, u32)> {
    let (sr, er) = if start_row <= end_row {
        (start_row, end_row)
    } else {
        (end_row, start_row)
    };
    let (sc, ec) = if start_col <= end_col {
        (start_col, end_col)
    } else {
        (end_col, start_col)
    };
    if sr == er && sc == ec {
        Some((sheet.to_string(), sr, sc))
    } else {
        None
    }
}

pub(crate) fn classify_reference(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<ReferenceKind> {
    match expr {
        Expr::Reference { sheet, row, col } => Ok(ReferenceKind::Single {
            sheet: sheet.clone(),
            row: *row,
            col: *col,
        }),
        Expr::Range(range) => {
            if let Some((sheet, row, col)) = single_cell_from_range(
                range.sheet.as_str(),
                range.start_row,
                range.end_row,
                range.start_col,
                range.end_col,
            ) {
                Ok(ReferenceKind::Single { sheet, row, col })
            } else {
                Ok(ReferenceKind::Range)
            }
        },
        Expr::Name(name) => match ctx.resolve_name(current_sheet, name.as_str())? {
            Some(ResolvedName::Cell { sheet, row, col }) => {
                Ok(ReferenceKind::Single { sheet, row, col })
            },
            Some(ResolvedName::Range(range)) => {
                if let Some((sheet, row, col)) = single_cell_from_range(
                    range.sheet.as_str(),
                    range.start_row,
                    range.end_row,
                    range.start_col,
                    range.end_col,
                ) {
                    Ok(ReferenceKind::Single { sheet, row, col })
                } else {
                    Ok(ReferenceKind::Range)
                }
            },
            None => Ok(ReferenceKind::Error(CellValue::Error(format!(
                "Unknown name: {}",
                name
            )))),
        },
        _ => Ok(ReferenceKind::None),
    }
}

pub(crate) async fn eval_isblank(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISBLANK expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(is_blank(&v)))
}

pub(crate) async fn eval_iserror(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISERROR expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(error_code(&v).is_some()))
}

pub(crate) async fn eval_iserr(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISERR expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let result = error_code(&v).is_some_and(|code| !is_na_error(code));
    Ok(CellValue::Bool(result))
}

pub(crate) async fn eval_isna(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISNA expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let result = error_code(&v).is_some_and(is_na_error);
    Ok(CellValue::Bool(result))
}

pub(crate) async fn eval_isnumber(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISNUMBER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let result = matches!(
        v,
        CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_)
    );
    Ok(CellValue::Bool(result))
}

pub(crate) async fn eval_istext(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISTEXT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(matches!(v, CellValue::String(_))))
}

pub(crate) async fn eval_islogical(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISLOGICAL expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(matches!(v, CellValue::Bool(_))))
}

pub(crate) async fn eval_isnontext(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISNONTEXT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(!matches!(v, CellValue::String(_))))
}

pub(crate) async fn eval_iseven(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISEVEN expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let num = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ISEVEN expects a numeric argument".to_string(),
            ));
        },
    };
    let truncated = num.trunc();
    Ok(CellValue::Bool(is_even_number(truncated)))
}

pub(crate) async fn eval_isodd(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISODD expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let num = match to_number(&v) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "ISODD expects a numeric argument".to_string(),
            ));
        },
    };
    let truncated = num.trunc();
    Ok(CellValue::Bool(!is_even_number(truncated)))
}

pub(crate) async fn eval_isformula(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISFORMULA expects 1 argument".to_string()));
    }
    match classify_reference(ctx, current_sheet, &args[0])? {
        ReferenceKind::Single { sheet, row, col } => {
            let raw = ctx.raw_cell_value(sheet.as_str(), row, col).await?;
            Ok(CellValue::Bool(matches!(raw, CellValue::Formula { .. })))
        },
        ReferenceKind::Range => Ok(CellValue::Error(
            "ISFORMULA expects a single cell reference".to_string(),
        )),
        ReferenceKind::None => Ok(CellValue::Error(
            "ISFORMULA expects a cell reference".to_string(),
        )),
        ReferenceKind::Error(err) => Ok(err),
    }
}

pub(crate) async fn eval_formulatext(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "FORMULATEXT expects 0 or 1 argument".to_string(),
        ));
    }

    let target = if args.is_empty() {
        if let Some((sheet, row, col)) = ctx.current_position() {
            Some((sheet, row, col))
        } else {
            None
        }
    } else {
        match classify_reference(ctx, current_sheet, &args[0])? {
            ReferenceKind::Single { sheet, row, col } => Some((sheet, row, col)),
            ReferenceKind::Range => None,
            ReferenceKind::None => None,
            ReferenceKind::Error(err) => return Ok(err),
        }
    };

    let (sheet, row, col) = match target {
        Some(t) => t,
        None => return Ok(CellValue::Error("#N/A".to_string())),
    };

    let raw = ctx.raw_cell_value(sheet.as_str(), row, col).await?;
    match raw {
        CellValue::Formula { formula, .. } => {
            let mut result = String::from("=");
            result.push_str(&formula);
            Ok(CellValue::String(result))
        },
        _ => Ok(CellValue::Error("#N/A".to_string())),
    }
}

pub(crate) async fn eval_info(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("INFO expects 1 argument".to_string()));
    }
    let info_type = evaluate_expression(ctx, current_sheet, &args[0]).await?;

    let info_type = match info_type {
        CellValue::String(s) => s.trim().to_ascii_lowercase(),
        _ => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    if info_type.is_empty() {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }

    let result = match info_type.as_str() {
        "recalc" => CellValue::String("Automatic".to_string()),
        "system" => CellValue::String("pcdos".to_string()),
        _ => CellValue::Error("#N/A".to_string()),
    };

    Ok(result)
}

pub(crate) async fn eval_isref(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ISREF expects 1 argument".to_string()));
    }
    match classify_reference(ctx, current_sheet, &args[0])? {
        ReferenceKind::Single { .. } | ReferenceKind::Range => Ok(CellValue::Bool(true)),
        ReferenceKind::None => Ok(CellValue::Bool(false)),
        ReferenceKind::Error(err) => Ok(err),
    }
}

pub(crate) async fn eval_na(_: EvalCtx<'_>, _: &str, args: &[Expr]) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("NA expects no arguments".to_string()));
    }
    Ok(CellValue::Error("#N/A".to_string()))
}

pub(crate) async fn eval_iferror(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("IFERROR expects 2 arguments".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if error_code(&value).is_some() {
        evaluate_expression(ctx, current_sheet, &args[1]).await
    } else {
        Ok(value)
    }
}

pub(crate) async fn eval_ifna(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("IFNA expects 2 arguments".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if error_code(&value).is_some_and(is_na_error) {
        evaluate_expression(ctx, current_sheet, &args[1]).await
    } else {
        Ok(value)
    }
}

pub(crate) async fn eval_n(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("N expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let result = match v {
        CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_) => v,
        CellValue::Bool(true) => CellValue::Int(1),
        CellValue::Bool(false) => CellValue::Int(0),
        CellValue::Error(_) => v,
        CellValue::Empty => CellValue::Int(0),
        CellValue::String(_) => CellValue::Int(0),
        CellValue::Formula { .. } => CellValue::Int(0),
    };
    Ok(result)
}

pub(crate) async fn eval_t(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("T expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match v {
        CellValue::String(_) => Ok(v),
        CellValue::Error(_) => Ok(v),
        _ => Ok(CellValue::String(String::new())),
    }
}

pub(crate) async fn eval_type(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("TYPE expects 1 argument".to_string()));
    }
    if matches!(args[0], Expr::Range(_)) {
        return Ok(CellValue::Int(64));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let code = match v {
        CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_) => 1,
        CellValue::String(_) => 2,
        CellValue::Bool(_) => 4,
        CellValue::Error(_) => return Ok(v),
        CellValue::Empty => 1,
        CellValue::Formula { .. } => 64,
    };
    Ok(CellValue::Int(code))
}

pub(crate) async fn eval_value(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("VALUE expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match v {
        CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_) => Ok(v),
        CellValue::Error(_) => Ok(v),
        CellValue::String(s) => {
            let trimmed = s.trim();
            if trimmed.is_empty() {
                return Ok(CellValue::Error("#VALUE!".to_string()));
            }
            // For now, simple float parsing
            if let Ok(f) = trimmed.parse::<f64>() {
                Ok(CellValue::Float(f))
            } else {
                Ok(CellValue::Error("#VALUE!".to_string()))
            }
        },
        CellValue::Bool(_) | CellValue::Empty | CellValue::Formula { .. } => {
            Ok(CellValue::Error("#VALUE!".to_string()))
        },
    }
}

pub(crate) async fn eval_sheet(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "SHEET expects 0 or 1 argument".to_string(),
        ));
    }

    let target_sheet = if args.is_empty() {
        current_sheet.to_string()
    } else {
        match classify_reference(ctx, current_sheet, &args[0])? {
            ReferenceKind::Single { sheet, .. } => sheet,
            ReferenceKind::Range => {
                // For a range, SHEET returns the index of the first sheet in the range.
                // In our current implementation, ranges are always on a single sheet.
                match &args[0] {
                    Expr::Range(r) => r.sheet.clone(),
                    _ => return Ok(CellValue::Error("#VALUE!".to_string())),
                }
            },
            ReferenceKind::None => return Ok(CellValue::Error("#VALUE!".to_string())),
            ReferenceKind::Error(err) => return Ok(err),
        }
    };

    // We need a way to get the sheet index from the context.
    // Since EngineCtx doesn't have it, we might need to add it or use a workaround.
    // For now, let's assume we can add it to EngineCtx.
    Ok(CellValue::Int(
        ctx.get_sheet_index(&target_sheet)
            .map(|idx| (idx + 1) as i64)
            .unwrap_or(0),
    ))
}

pub(crate) async fn eval_sheets(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() > 1 {
        return Ok(CellValue::Error(
            "SHEETS expects 0 or 1 argument".to_string(),
        ));
    }

    if args.is_empty() {
        return Ok(CellValue::Int(ctx.get_sheet_count() as i64));
    }

    match classify_reference(ctx, current_sheet, &args[0])? {
        ReferenceKind::Single { .. } | ReferenceKind::Range => {
            // In our current implementation, ranges/refs are always on a single sheet.
            // Excel supports multi-sheet references (3D references), but we don't yet.
            Ok(CellValue::Int(1))
        },
        ReferenceKind::None => Ok(CellValue::Int(ctx.get_sheet_count() as i64)),
        ReferenceKind::Error(err) => Ok(err),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    fn bool_expr(b: bool) -> Expr {
        Expr::Literal(CellValue::Bool(b))
    }

    #[test]
    fn test_error_code_with_error() {
        let err = CellValue::Error("#VALUE!".to_string());
        assert_eq!(error_code(&err), Some("#VALUE!"));
    }

    #[test]
    fn test_error_code_without_error() {
        let val = CellValue::Int(42);
        assert_eq!(error_code(&val), None);
    }

    #[test]
    fn test_is_na_error_true() {
        assert!(is_na_error("#N/A"));
        assert!(is_na_error("#n/a"));
    }

    #[test]
    fn test_is_na_error_false() {
        assert!(!is_na_error("#VALUE!"));
        assert!(!is_na_error("#REF!"));
    }

    #[test]
    fn test_is_even_number() {
        assert!(is_even_number(4.0));
        assert!(is_even_number(-2.0));
        assert!(is_even_number(0.0));
        assert!(!is_even_number(3.0));
        assert!(!is_even_number(-1.0));
    }

    #[tokio::test]
    async fn test_eval_isblank_with_blank() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Empty)];
        let result = eval_isblank(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isblank_with_value() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_isblank(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iserror_with_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#VALUE!".to_string()))];
        let result = eval_iserror(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iserror_with_value() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_iserror(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iserr_with_value_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#VALUE!".to_string()))];
        let result = eval_iserr(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iserr_with_na_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#N/A".to_string()))];
        let result = eval_iserr(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false) for #N/A"),
        }
    }

    #[tokio::test]
    async fn test_eval_isna_with_na() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#N/A".to_string()))];
        let result = eval_isna(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isna_with_other_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#VALUE!".to_string()))];
        let result = eval_isna(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isnumber_with_int() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_isnumber(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isnumber_with_float() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(std::f64::consts::PI)];
        let result = eval_isnumber(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isnumber_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_isnumber(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_istext_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_istext(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_istext_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_istext(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_islogical_with_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_islogical(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_islogical_with_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false)];
        let result = eval_islogical(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_islogical_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_islogical(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isnontext_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_isnontext(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isnontext_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_isnontext(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iseven_with_even() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(4.0)];
        let result = eval_iseven(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iseven_with_odd() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.0)];
        let result = eval_iseven(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iseven_with_non_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_iseven(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects a numeric")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_isodd_with_odd() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(3.0)];
        let result = eval_isodd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(true) => {},
            _ => panic!("Expected Bool(true)"),
        }
    }

    #[tokio::test]
    async fn test_eval_isodd_with_even() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(4.0)];
        let result = eval_isodd(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Bool(false) => {},
            _ => panic!("Expected Bool(false)"),
        }
    }

    #[tokio::test]
    async fn test_eval_na() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_na(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#N/A"),
            _ => panic!("Expected Error(#N/A)"),
        }
    }

    #[tokio::test]
    async fn test_eval_na_with_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0)];
        let result = eval_na(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects no arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_iferror_with_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::Error("#VALUE!".to_string())),
            num_expr(42.0),
        ];
        let result = eval_iferror(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(42) => {},
            _ => panic!("Expected Int(42)"),
        }
    }

    #[tokio::test]
    async fn test_eval_iferror_with_no_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(100.0), num_expr(42.0)];
        let result = eval_iferror(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(100.0) | CellValue::Int(100) => {},
            _ => panic!("Expected 100, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_ifna_with_na() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::Error("#N/A".to_string())),
            str_expr("not available"),
        ];
        let result = eval_ifna(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "not available"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_ifna_with_other_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::Error("#VALUE!".to_string())),
            str_expr("not available"),
        ];
        let result = eval_ifna(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected Error(#VALUE!)"),
        }
    }

    #[tokio::test]
    async fn test_eval_n_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_n(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(42.0) | CellValue::Int(42) => {},
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_eval_n_with_true() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_n(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(1) => {},
            _ => panic!("Expected Int(1)"),
        }
    }

    #[tokio::test]
    async fn test_eval_n_with_false() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(false)];
        let result = eval_n(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(0) => {},
            _ => panic!("Expected Int(0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_n_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_n(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(0) => {},
            _ => panic!("Expected Int(0)"),
        }
    }

    #[tokio::test]
    async fn test_eval_t_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_t(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "hello"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_t_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_t(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert!(s.is_empty()),
            _ => panic!("Expected empty String"),
        }
    }

    #[tokio::test]
    async fn test_eval_type_with_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_type(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(1) => {},
            _ => panic!("Expected Int(1)"),
        }
    }

    #[tokio::test]
    async fn test_eval_type_with_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_type(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(2) => {},
            _ => panic!("Expected Int(2)"),
        }
    }

    #[tokio::test]
    async fn test_eval_type_with_bool() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![bool_expr(true)];
        let result = eval_type(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(4) => {},
            _ => panic!("Expected Int(4)"),
        }
    }

    #[tokio::test]
    async fn test_eval_type_with_error() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Error("#VALUE!".to_string()))];
        let result = eval_type(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_value_with_number_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("123.45")];
        let result = eval_value(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 123.45).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_value_with_non_number() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("hello")];
        let result = eval_value(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_value_with_empty() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("")];
        let result = eval_value(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_info_recalc() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("recalc")];
        let result = eval_info(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "Automatic"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_info_system() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("system")];
        let result = eval_info(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::String(s) => assert_eq!(s, "pcdos"),
            _ => panic!("Expected String"),
        }
    }

    #[tokio::test]
    async fn test_eval_info_invalid() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("invalid")];
        let result = eval_info(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#N/A")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_info_non_string() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_info(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_sheets_no_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_sheets(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(n) => assert!(n >= 0),
            _ => panic!("Expected Int"),
        }
    }

    #[tokio::test]
    async fn test_eval_sheet_no_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_sheet(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Int(n) => assert!(n >= 0),
            _ => panic!("Expected Int"),
        }
    }
}
