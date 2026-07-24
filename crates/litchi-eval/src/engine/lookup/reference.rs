use crate::parser::Expr;
use litchi_core::sheet::{CellValue, Result};

use super::super::info::{ReferenceKind, classify_reference};
use super::super::{EvalCtx, evaluate_expression, to_bool, to_number, to_text};

fn column_index_to_letters(col: u32) -> String {
    let mut n = col;
    let mut letters = Vec::new();
    while n > 0 {
        n -= 1;
        letters.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    letters.iter().rev().collect()
}

/// Quotes a sheet name the way Excel does in references: single quotes are
/// doubled, and the name is wrapped in quotes when it is not a "plain" name
/// (letters/digits/underscore/period, not starting with a digit).
fn format_sheet_prefix(sheet: &str) -> String {
    let mut chars = sheet.chars();
    let plain = chars
        .next()
        .is_some_and(|c| c.is_alphabetic() || c == '_' || c == '\\')
        && sheet
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.');
    if plain {
        format!("{}!", sheet)
    } else {
        format!("'{}'!", sheet.replace('\'', "''"))
    }
}

/// ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]) — builds a
/// cell reference as text, in A1 or R1C1 style, optionally sheet-qualified.
pub(crate) async fn eval_address(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 5 {
        return Ok(CellValue::Error(
            "ADDRESS expects 2 to 5 arguments (row_num, column_num, [abs_num], [a1], [sheet_text])"
                .to_string(),
        ));
    }

    let row_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let row = match to_number(&row_val) {
        Some(n) if n.trunc() >= 1.0 && n.trunc() <= u32::MAX as f64 => n.trunc() as u32,
        Some(_) => return Ok(CellValue::Error("#VALUE!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let col_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let col = match to_number(&col_val) {
        Some(n) if n.trunc() >= 1.0 && n.trunc() <= u32::MAX as f64 => n.trunc() as u32,
        Some(_) => return Ok(CellValue::Error("#VALUE!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let abs_num = if args.len() >= 3 {
        let v = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&v) {
            Some(n) if (1.0..=4.0).contains(&n.trunc()) => n.trunc() as u8,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    } else {
        1
    };
    let (abs_row, abs_col) = match abs_num {
        1 => (true, true),
        2 => (true, false),
        3 => (false, true),
        _ => (false, false),
    };

    let a1 = if args.len() >= 4 {
        let v = evaluate_expression(ctx, current_sheet, &args[3]).await?;
        to_bool(&v)
    } else {
        true
    };

    let reference = if a1 {
        format!(
            "{}{}{}{}",
            if abs_col { "$" } else { "" },
            column_index_to_letters(col),
            if abs_row { "$" } else { "" },
            row
        )
    } else {
        // R1C1 style: absolute positions are plain, relative ones bracketed.
        format!(
            "{}{}",
            if abs_row {
                format!("R{}", row)
            } else {
                format!("R[{}]", row)
            },
            if abs_col {
                format!("C{}", col)
            } else {
                format!("C[{}]", col)
            }
        )
    };

    let text = if args.len() == 5 {
        let v = evaluate_expression(ctx, current_sheet, &args[4]).await?;
        if let CellValue::Error(_) = v {
            return Ok(v);
        }
        let sheet = to_text(&v);
        if sheet.is_empty() {
            reference
        } else {
            format!("{}{}", format_sheet_prefix(&sheet), reference)
        }
    } else {
        reference
    };

    Ok(CellValue::String(text))
}

/// AREAS(reference) — the number of areas in a reference.
///
/// Limitation: the parser's reference model (`Expr::Reference`,
/// `Expr::Range`, resolvable names) can only express a single contiguous
/// area; multi-area union references such as `(A1:B2,D4:E5)` have no AST
/// representation, so every valid reference here counts as exactly one area.
pub(crate) async fn eval_areas(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("AREAS expects 1 argument".to_string()));
    }

    match classify_reference(ctx, current_sheet, &args[0])? {
        ReferenceKind::Single { .. } | ReferenceKind::Range => Ok(CellValue::Int(1)),
        ReferenceKind::None => Ok(CellValue::Error("#VALUE!".to_string())),
        ReferenceKind::Error(v) => Ok(v),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engine::test_helpers::TestEngine;
    use crate::parser::ast::RangeRef;

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

    async fn address(args: Vec<Expr>) -> CellValue {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        eval_address(ctx, "Sheet1", &args).await.unwrap()
    }

    #[tokio::test]
    async fn test_address_default_absolute() {
        let result = address(vec![num_expr(2.0), num_expr(3.0)]).await;
        assert_eq!(result, CellValue::String("$C$2".to_string()));
    }

    #[tokio::test]
    async fn test_address_abs_num_variants() {
        assert_eq!(
            address(vec![num_expr(2.0), num_expr(3.0), num_expr(1.0)]).await,
            CellValue::String("$C$2".to_string())
        );
        assert_eq!(
            address(vec![num_expr(2.0), num_expr(3.0), num_expr(2.0)]).await,
            CellValue::String("C$2".to_string())
        );
        assert_eq!(
            address(vec![num_expr(2.0), num_expr(3.0), num_expr(3.0)]).await,
            CellValue::String("$C2".to_string())
        );
        assert_eq!(
            address(vec![num_expr(2.0), num_expr(3.0), num_expr(4.0)]).await,
            CellValue::String("C2".to_string())
        );
    }

    #[tokio::test]
    async fn test_address_r1c1_style() {
        assert_eq!(
            address(vec![
                num_expr(2.0),
                num_expr(3.0),
                num_expr(1.0),
                bool_expr(false)
            ])
            .await,
            CellValue::String("R2C3".to_string())
        );
        assert_eq!(
            address(vec![
                num_expr(2.0),
                num_expr(3.0),
                num_expr(2.0),
                bool_expr(false)
            ])
            .await,
            CellValue::String("R2C[3]".to_string())
        );
        assert_eq!(
            address(vec![
                num_expr(2.0),
                num_expr(3.0),
                num_expr(4.0),
                bool_expr(false)
            ])
            .await,
            CellValue::String("R[2]C[3]".to_string())
        );
    }

    #[tokio::test]
    async fn test_address_multi_letter_column() {
        // Column 28 -> AB
        let result = address(vec![num_expr(1.0), num_expr(28.0)]).await;
        assert_eq!(result, CellValue::String("$AB$1".to_string()));
    }

    #[tokio::test]
    async fn test_address_invalid_abs_num() {
        let result = address(vec![num_expr(1.0), num_expr(1.0), num_expr(5.0)]).await;
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected #VALUE!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_address_invalid_row_or_col() {
        for args in [
            vec![num_expr(0.0), num_expr(1.0)],
            vec![num_expr(1.0), num_expr(0.0)],
            vec![num_expr(-3.0), num_expr(1.0)],
        ] {
            let result = address(args).await;
            match result {
                CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
                _ => panic!("Expected #VALUE!, got {:?}", result),
            }
        }
    }

    #[tokio::test]
    async fn test_address_non_numeric() {
        let result = address(vec![str_expr("x"), num_expr(1.0)]).await;
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected #VALUE!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_address_plain_sheet_name_unquoted() {
        let result = address(vec![
            num_expr(1.0),
            num_expr(1.0),
            num_expr(1.0),
            bool_expr(true),
            str_expr("Sheet2"),
        ])
        .await;
        assert_eq!(result, CellValue::String("Sheet2!$A$1".to_string()));
    }

    #[tokio::test]
    async fn test_address_sheet_name_with_space_quoted() {
        let result = address(vec![
            num_expr(1.0),
            num_expr(1.0),
            num_expr(1.0),
            bool_expr(true),
            str_expr("My Sheet"),
        ])
        .await;
        assert_eq!(result, CellValue::String("'My Sheet'!$A$1".to_string()));
    }

    #[tokio::test]
    async fn test_address_sheet_name_quote_escaping() {
        let result = address(vec![
            num_expr(1.0),
            num_expr(1.0),
            num_expr(1.0),
            bool_expr(true),
            str_expr("O'Brien"),
        ])
        .await;
        assert_eq!(result, CellValue::String("'O''Brien'!$A$1".to_string()));
    }

    #[tokio::test]
    async fn test_address_wrong_arg_count() {
        let result = address(vec![num_expr(1.0)]).await;
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 to 5")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_areas_single_cell_reference() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Reference {
            sheet: "Sheet1".to_string(),
            row: 1,
            col: 1,
        }];
        let result = eval_areas(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(1));
    }

    #[tokio::test]
    async fn test_areas_range_reference() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Range(RangeRef {
            sheet: "Sheet1".to_string(),
            start_row: 1,
            start_col: 1,
            end_row: 5,
            end_col: 3,
        })];
        let result = eval_areas(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(1));
    }

    #[tokio::test]
    async fn test_areas_non_reference() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(42.0)];
        let result = eval_areas(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            _ => panic!("Expected #VALUE!, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_areas_wrong_arg_count() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_areas(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }
}
