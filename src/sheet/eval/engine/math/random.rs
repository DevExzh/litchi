use rand::Rng;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_rand(
    _ctx: EvalCtx<'_>,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("RAND expects 0 arguments".to_string()));
    }
    let mut rng = rand::rng();
    Ok(CellValue::Float(rng.random::<f64>()))
}

pub(crate) async fn eval_randbetween(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "RANDBETWEEN expects 2 arguments (bottom, top)".to_string(),
        ));
    }
    let bottom_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let top_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let bottom = match to_number(&bottom_val) {
        Some(n) => n.trunc() as i64,
        None => {
            return Ok(CellValue::Error(
                "RANDBETWEEN bottom must be numeric".to_string(),
            ));
        },
    };
    let top = match to_number(&top_val) {
        Some(n) => n.trunc() as i64,
        None => {
            return Ok(CellValue::Error(
                "RANDBETWEEN top must be numeric".to_string(),
            ));
        },
    };
    if bottom > top {
        return Ok(CellValue::Error(
            "RANDBETWEEN bottom must be less than or equal to top".to_string(),
        ));
    }
    let mut rng = rand::rng();
    let value = if bottom == top {
        bottom
    } else {
        rng.random_range(bottom..=top)
    };
    Ok(CellValue::Int(value))
}
