use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_arabic(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ARABIC expects 1 argument".to_string()));
    }
    let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?)
        .trim()
        .to_uppercase();
    if text.is_empty() {
        return Ok(CellValue::Int(0));
    }

    let mut result = 0i64;
    let mut last_val = 0i64;

    // Handle negative sign
    let (haystack, multiplier) = if let Some(stripped) = text.strip_prefix('-') {
        (stripped, -1)
    } else {
        (text.as_str(), 1)
    };

    for c in haystack.chars().rev() {
        let val = match c {
            'I' => 1,
            'V' => 5,
            'X' => 10,
            'L' => 50,
            'C' => 100,
            'D' => 500,
            'M' => 1000,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        if val < last_val {
            result -= val;
        } else {
            result += val;
        }
        last_val = val;
    }

    Ok(CellValue::Int(result * multiplier))
}

pub(crate) async fn eval_roman(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "ROMAN expects 1 or 2 arguments".to_string(),
        ));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let num = match crate::sheet::eval::engine::to_number(&val) {
        Some(n) => n.trunc() as i64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if !(0..=3999).contains(&num) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    if num == 0 {
        return Ok(CellValue::String(String::new()));
    }

    // Excel ROMAN function has different forms (0 to 4), but usually 0 (classic) is used.
    // We'll implement classic (form 0) for now.
    let mut n = num;
    let mut result = String::new();
    let mapping = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    for (val, sym) in mapping {
        while n >= val {
            result.push_str(sym);
            n -= val;
        }
    }

    Ok(CellValue::String(result))
}
