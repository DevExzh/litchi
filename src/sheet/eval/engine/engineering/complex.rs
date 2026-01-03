use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use num_complex::Complex64;

/// Parses an Excel complex number string (e.g., "3+4i", "j", "-2i") into a Complex64.
fn parse_complex(s: &str) -> Option<Complex64> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    // Determine the suffix (i or j)
    let _suffix = if s.ends_with('i') {
        'i'
    } else if s.ends_with('j') {
        'j'
    } else {
        // Pure real number
        return s.parse::<f64>().ok().map(|re| Complex64::new(re, 0.0));
    };

    let s_no_suffix = &s[..s.len() - 1];

    // Case: just "i" or "j" or "-i" or "-j" or "+i" or "+j"
    if s_no_suffix.is_empty() {
        return Some(Complex64::new(0.0, 1.0));
    }
    if s_no_suffix == "-" {
        return Some(Complex64::new(0.0, -1.0));
    }
    if s_no_suffix == "+" {
        return Some(Complex64::new(0.0, 1.0));
    }

    // Try parsing as pure imaginary first (e.g., "4i", "-2j")
    if let Ok(im) = s_no_suffix.parse::<f64>() {
        return Some(Complex64::new(0.0, im));
    }

    // Otherwise, it must be re+im[suffix] or re-im[suffix]
    // We need to find the last '+' or '-' that isn't at the very beginning or after 'e' (scientific notation)
    let bytes = s_no_suffix.as_bytes();
    let mut split_pos = None;
    for i in (1..bytes.len()).rev() {
        let b = bytes[i];
        if (b == b'+' || b == b'-') && bytes[i - 1] != b'e' && bytes[i - 1] != b'E' {
            split_pos = Some(i);
            break;
        }
    }

    if let Some(pos) = split_pos {
        let re_part = &s_no_suffix[..pos];
        let im_part = &s_no_suffix[pos..];

        let re = re_part.parse::<f64>().ok()?;
        let im = if im_part == "+" {
            1.0
        } else if im_part == "-" {
            -1.0
        } else {
            im_part.parse::<f64>().ok()?
        };

        return Some(Complex64::new(re, im));
    }

    None
}

/// Formats a Complex64 into an Excel complex number string.
fn format_complex(c: Complex64, suffix: &str) -> String {
    let re = c.re;
    let im = c.im;

    if im == 0.0 {
        return re.to_string();
    }

    let im_str = if im == 1.0 {
        suffix.to_string()
    } else if im == -1.0 {
        format!("-{}", suffix)
    } else {
        format!("{}{}", im, suffix)
    };

    if re == 0.0 {
        im_str
    } else if im > 0.0 {
        format!("{}+{}", re, im_str)
    } else {
        format!("{}{}", re, im_str)
    }
}

pub(crate) async fn eval_complex(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "COMPLEX expects 2 or 3 arguments".to_string(),
        ));
    }

    let re_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let im_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;

    let re = match to_number(&re_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let im = match to_number(&im_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let suffix = if args.len() == 3 {
        let s_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        let s = to_text(&s_val);
        if s != "i" && s != "j" {
            return Ok(CellValue::Error("#VALUE!".to_string()));
        }
        s
    } else {
        "i".to_string()
    };

    Ok(CellValue::String(format_complex(
        Complex64::new(re, im),
        &suffix,
    )))
}

pub(crate) async fn eval_imabs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("IMABS expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::Float(c.norm())),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imaginary(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("IMAGINARY expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::Float(c.im)),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imreal(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("IMREAL expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::Float(c.re)),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imargument(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "IMARGUMENT expects 1 argument".to_string(),
        ));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::Float(c.arg())),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imconjugate(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "IMCONJUGATE expects 1 argument".to_string(),
        ));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    let suffix = if text.ends_with('j') { "j" } else { "i" };
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::String(format_complex(c.conj(), suffix))),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imsum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "IMSUM expects at least 1 argument".to_string(),
        ));
    }
    let mut total = Complex64::new(0.0, 0.0);
    let mut suffix = "i";

    for arg in args {
        let val = evaluate_expression(ctx, current_sheet, arg).await?;
        let text = to_text(&val);
        if text.ends_with('j') {
            suffix = "j";
        }
        match parse_complex(&text) {
            Some(c) => total += c,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    }
    Ok(CellValue::String(format_complex(total, suffix)))
}

pub(crate) async fn eval_imsub(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("IMSUB expects 2 arguments".to_string()));
    }
    let val1 = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let val2 = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let text1 = to_text(&val1);
    let text2 = to_text(&val2);

    let suffix = if text1.ends_with('j') || text2.ends_with('j') {
        "j"
    } else {
        "i"
    };

    match (parse_complex(&text1), parse_complex(&text2)) {
        (Some(c1), Some(c2)) => Ok(CellValue::String(format_complex(c1 - c2, suffix))),
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_improduct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "IMPRODUCT expects at least 1 argument".to_string(),
        ));
    }
    let mut total = Complex64::new(1.0, 0.0);
    let mut suffix = "i";

    for arg in args {
        let val = evaluate_expression(ctx, current_sheet, arg).await?;
        let text = to_text(&val);
        if text.ends_with('j') {
            suffix = "j";
        }
        match parse_complex(&text) {
            Some(c) => total *= c,
            None => return Ok(CellValue::Error("#VALUE!".to_string())),
        }
    }
    Ok(CellValue::String(format_complex(total, suffix)))
}

pub(crate) async fn eval_imdiv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("IMDIV expects 2 arguments".to_string()));
    }
    let val1 = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let val2 = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let text1 = to_text(&val1);
    let text2 = to_text(&val2);

    let suffix = if text1.ends_with('j') || text2.ends_with('j') {
        "j"
    } else {
        "i"
    };

    match (parse_complex(&text1), parse_complex(&text2)) {
        (Some(c1), Some(c2)) => {
            if c2.re == 0.0 && c2.im == 0.0 {
                return Ok(CellValue::Error("#NUM!".to_string()));
            }
            Ok(CellValue::String(format_complex(c1 / c2, suffix)))
        },
        _ => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

pub(crate) async fn eval_imsin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMSIN", |c| c.sin()).await
}

pub(crate) async fn eval_imcos(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMCOS", |c| c.cos()).await
}

pub(crate) async fn eval_imtan(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMTAN", |c| c.tan()).await
}

pub(crate) async fn eval_imsinh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMSINH", |c| c.sinh()).await
}

pub(crate) async fn eval_imcosh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMCOSH", |c| c.cosh()).await
}

pub(crate) async fn eval_imtanh(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMTANH", |c| c.tanh()).await
}

pub(crate) async fn eval_imcsc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMCSC", |c| {
        let s = c.sin();
        Complex64::new(1.0, 0.0) / s
    })
    .await
}

pub(crate) async fn eval_imcsch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMCSCH", |c| {
        let s = c.sinh();
        Complex64::new(1.0, 0.0) / s
    })
    .await
}

pub(crate) async fn eval_imsec(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMSEC", |c| {
        let s = c.cos();
        Complex64::new(1.0, 0.0) / s
    })
    .await
}

pub(crate) async fn eval_imsech(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMSECH", |c| {
        let s = c.cosh();
        Complex64::new(1.0, 0.0) / s
    })
    .await
}

pub(crate) async fn eval_imcot(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMCOT", |c| {
        let t = c.tan();
        Complex64::new(1.0, 0.0) / t
    })
    .await
}

pub(crate) async fn eval_imsqrt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMSQRT", |c| c.sqrt()).await
}

pub(crate) async fn eval_imln(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMLN", |c| c.ln()).await
}

pub(crate) async fn eval_imlog10(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMLOG10", |c| c.log(10.0)).await
}

pub(crate) async fn eval_imlog2(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMLOG2", |c| c.log(2.0)).await
}

pub(crate) async fn eval_imexp(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_unary_complex(ctx, current_sheet, args, "IMEXP", |c| c.exp()).await
}

pub(crate) async fn eval_impower(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("IMPOWER expects 2 arguments".to_string()));
    }
    let val1 = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let val2 = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let text1 = to_text(&val1);
    let power = match to_number(&val2) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let suffix = if text1.ends_with('j') { "j" } else { "i" };

    match parse_complex(&text1) {
        Some(c1) => Ok(CellValue::String(format_complex(c1.powf(power), suffix))),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}

async fn eval_unary_complex<F>(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    name: &str,
    f: F,
) -> Result<CellValue>
where
    F: Fn(Complex64) -> Complex64,
{
    if args.len() != 1 {
        return Ok(CellValue::Error(format!("{} expects 1 argument", name)));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let text = to_text(&val);
    let suffix = if text.ends_with('j') { "j" } else { "i" };
    match parse_complex(&text) {
        Some(c) => Ok(CellValue::String(format_complex(f(c), suffix))),
        None => Ok(CellValue::Error("#VALUE!".to_string())),
    }
}
