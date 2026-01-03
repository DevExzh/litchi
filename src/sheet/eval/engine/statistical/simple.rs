use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use std::cmp::Ordering;

pub(crate) async fn eval_median(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "MEDIAN expects at least 1 argument".to_string(),
        ));
    }

    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    if numbers.is_empty() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let len = numbers.len();
    if len % 2 == 1 {
        Ok(CellValue::Float(numbers[len / 2]))
    } else {
        Ok(CellValue::Float(
            (numbers[len / 2 - 1] + numbers[len / 2]) / 2.0,
        ))
    }
}

pub(crate) async fn eval_mode_sngl(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "MODE.SNGL expects at least 1 argument".to_string(),
        ));
    }

    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    if numbers.is_empty() {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    // Sort to group duplicates
    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let mut max_count = 0;
    let mut current_count = 0;
    let mut current_val = numbers[0];
    let mut mode_val = None;

    for &n in &numbers {
        if (n - current_val).abs() < 1e-12 {
            current_count += 1;
        } else {
            if current_count > max_count {
                max_count = current_count;
                mode_val = Some(current_val);
            }
            current_val = n;
            current_count = 1;
        }
    }

    // Final check
    if current_count > max_count {
        max_count = current_count;
        mode_val = Some(current_val);
    }

    if max_count <= 1 {
        // No duplicates found
        Ok(CellValue::Error("#N/A".to_string()))
    } else {
        Ok(CellValue::Float(mode_val.unwrap()))
    }
}

pub(crate) async fn eval_stdev_s(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match eval_variance(ctx, current_sheet, args, true).await? {
        CellValue::Float(v) => Ok(CellValue::Float(v.sqrt())),
        other => Ok(other),
    }
}

pub(crate) async fn eval_stdev_p(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match eval_variance(ctx, current_sheet, args, false).await? {
        CellValue::Float(v) => Ok(CellValue::Float(v.sqrt())),
        other => Ok(other),
    }
}

pub(crate) async fn eval_var_s(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_variance(ctx, current_sheet, args, true).await
}

pub(crate) async fn eval_var_p(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_variance(ctx, current_sheet, args, false).await
}

pub(crate) async fn eval_geomean(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "GEOMEAN expects at least 1 argument".to_string(),
        ));
    }

    let mut numbers = Vec::new();
    for arg in args {
        let res = for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                if n <= 0.0 {
                    return Err(Box::new(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "positive numbers required",
                    )));
                }
                numbers.push(n);
            }
            Ok(())
        })
        .await;

        if res.is_err() {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
    }

    if numbers.is_empty() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let mut product = 1.0;
    let n = numbers.len() as f64;
    for x in numbers {
        product *= x.powf(1.0 / n);
    }

    Ok(CellValue::Float(product))
}

pub(crate) async fn eval_harmean(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "HARMEAN expects at least 1 argument".to_string(),
        ));
    }

    let mut numbers = Vec::new();
    for arg in args {
        let res = for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                if n <= 0.0 {
                    return Err(Box::new(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "positive numbers required",
                    )));
                }
                numbers.push(n);
            }
            Ok(())
        })
        .await;

        if res.is_err() {
            return Ok(CellValue::Error("#NUM!".to_string()));
        }
    }

    if numbers.is_empty() {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let mut sum_inv = 0.0;
    let n = numbers.len() as f64;
    for x in numbers {
        sum_inv += 1.0 / x;
    }

    Ok(CellValue::Float(n / sum_inv))
}

pub(crate) async fn eval_trimmean(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "TRIMMEAN expects 2 arguments (array, percent)".to_string(),
        ));
    }

    let mut numbers = Vec::new();
    for_each_value_in_expr(ctx, current_sheet, &args[0], |val| {
        if let Some(n) = to_number(val) {
            numbers.push(n);
        }
        Ok(())
    })
    .await?;

    let percent_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let percent = match to_number(&percent_val) {
        Some(p) if (0.0..1.0).contains(&p) => p,
        _ => return Ok(CellValue::Error("#NUM!".to_string())),
    };

    let n = numbers.len();
    if n == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let k = (n as f64 * percent).floor() as usize;
    let trim_count = k / 2; // k points are trimmed in total, k/2 from each end

    if trim_count * 2 >= n {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let trimmed = &numbers[trim_count..n - trim_count];
    let sum: f64 = trimmed.iter().sum();
    Ok(CellValue::Float(sum / trimmed.len() as f64))
}

pub(crate) async fn eval_skew(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    let n = numbers.len();
    if n < 3 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean = numbers.iter().sum::<f64>() / n as f64;
    let sum_sq_diff: f64 = numbers.iter().map(|&x| (x - mean).powi(2)).sum();
    let stdev = (sum_sq_diff / (n - 1) as f64).sqrt();

    if stdev == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let sum_cubed_diff: f64 = numbers.iter().map(|&x| ((x - mean) / stdev).powi(3)).sum();
    let skew = (n as f64 / ((n - 1) as f64 * (n - 2) as f64)) * sum_cubed_diff;

    Ok(CellValue::Float(skew))
}

pub(crate) async fn eval_skew_p(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    let n = numbers.len();
    if n == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean = numbers.iter().sum::<f64>() / n as f64;
    let sum_sq_diff: f64 = numbers.iter().map(|&x| (x - mean).powi(2)).sum();
    let stdev_p = (sum_sq_diff / n as f64).sqrt();

    if stdev_p == 0.0 {
        return Ok(CellValue::Float(0.0));
    }

    let sum_cubed_diff: f64 = numbers
        .iter()
        .map(|&x| ((x - mean) / stdev_p).powi(3))
        .sum();
    let skew_p = sum_cubed_diff / n as f64;

    Ok(CellValue::Float(skew_p))
}

pub(crate) async fn eval_kurt(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    let n = numbers.len();
    if n < 4 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean = numbers.iter().sum::<f64>() / n as f64;
    let sum_sq_diff: f64 = numbers.iter().map(|&x| (x - mean).powi(2)).sum();
    let stdev = (sum_sq_diff / (n - 1) as f64).sqrt();

    if stdev == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let sum_fourth_diff: f64 = numbers.iter().map(|&x| ((x - mean) / stdev).powi(4)).sum();

    let term1 = (n as f64 * (n + 1) as f64) / ((n - 1) as f64 * (n - 2) as f64 * (n - 3) as f64);
    let term2 = (3.0 * ((n - 1) as f64).powi(2)) / ((n - 2) as f64 * (n - 3) as f64);

    let kurt = term1 * sum_fourth_diff - term2;

    Ok(CellValue::Float(kurt))
}

pub(crate) async fn eval_stdev_a(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match eval_variance_a(ctx, current_sheet, args, true).await? {
        CellValue::Float(v) => Ok(CellValue::Float(v.sqrt())),
        other => Ok(other),
    }
}

pub(crate) async fn eval_stdev_pa(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match eval_variance_a(ctx, current_sheet, args, false).await? {
        CellValue::Float(v) => Ok(CellValue::Float(v.sqrt())),
        other => Ok(other),
    }
}

pub(crate) async fn eval_var_a(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_variance_a(ctx, current_sheet, args, true).await
}

pub(crate) async fn eval_var_pa(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_variance_a(ctx, current_sheet, args, false).await
}

async fn eval_variance_a(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    sample: bool,
) -> Result<CellValue> {
    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            match val {
                CellValue::Empty => {}, // Skip empty cells
                CellValue::Bool(b) => numbers.push(if *b { 1.0 } else { 0.0 }),
                CellValue::String(_) => numbers.push(0.0),
                CellValue::Error(e) => {
                    return Err(Box::new(std::io::Error::other(e.clone())));
                },
                _ => {
                    if let Some(n) = to_number(val) {
                        numbers.push(n);
                    }
                },
            }
            Ok(())
        })
        .await?;
    }

    let n = numbers.len();
    if n == 0 || (sample && n < 2) {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean = numbers.iter().sum::<f64>() / n as f64;
    let sum_sq_diff: f64 = numbers.iter().map(|&x| (x - mean).powi(2)).sum();

    let divisor = if sample { (n - 1) as f64 } else { n as f64 };
    Ok(CellValue::Float(sum_sq_diff / divisor))
}

async fn eval_variance(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    sample: bool,
) -> Result<CellValue> {
    let mut numbers = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |val| {
            if let Some(n) = to_number(val) {
                numbers.push(n);
            }
            Ok(())
        })
        .await?;
    }

    let n = numbers.len();
    if n == 0 || (sample && n < 2) {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean = numbers.iter().sum::<f64>() / n as f64;
    let sum_sq_diff: f64 = numbers.iter().map(|&x| (x - mean).powi(2)).sum();

    let divisor = if sample { (n - 1) as f64 } else { n as f64 };
    Ok(CellValue::Float(sum_sq_diff / divisor))
}

pub(crate) async fn eval_fisher(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("FISHER expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let x = match to_number(&val) {
        Some(n) if n > -1.0 && n < 1.0 => n,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float(0.5 * ((1.0 + x) / (1.0 - x)).ln()))
}

pub(crate) async fn eval_fisherinv(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("FISHERINV expects 1 argument".to_string()));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let y = match to_number(&val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let exp2y = (2.0 * y).exp();
    Ok(CellValue::Float((exp2y - 1.0) / (exp2y + 1.0)))
}

pub(crate) async fn eval_standardize(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "STANDARDIZE expects 3 arguments (x, mean, standard_dev)".to_string(),
        ));
    }
    let x = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let mean = match to_number(&evaluate_expression(ctx, current_sheet, &args[1]).await?) {
        Some(v) => v,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let stdev = match to_number(&evaluate_expression(ctx, current_sheet, &args[2]).await?) {
        Some(v) if v > 0.0 => v,
        Some(_) => return Ok(CellValue::Error("#NUM!".to_string())),
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    Ok(CellValue::Float((x - mean) / stdev))
}

pub(crate) async fn eval_covar_p(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_covariance(ctx, current_sheet, args, false).await
}

pub(crate) async fn eval_covar_s(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_covariance(ctx, current_sheet, args, true).await
}

pub(crate) async fn eval_correl(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("CORREL expects 2 arguments".to_string()));
    }

    let (xs, ys) = collect_aligned_numeric_pairs(ctx, current_sheet, &args[0], &args[1]).await?;
    let n = xs.len();
    if n == 0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean_x = xs.iter().sum::<f64>() / n as f64;
    let mean_y = ys.iter().sum::<f64>() / n as f64;

    let mut sum_prod_diff = 0.0;
    let mut sum_sq_diff_x = 0.0;
    let mut sum_sq_diff_y = 0.0;

    for i in 0..n {
        let dx = xs[i] - mean_x;
        let dy = ys[i] - mean_y;
        sum_prod_diff += dx * dy;
        sum_sq_diff_x += dx * dx;
        sum_sq_diff_y += dy * dy;
    }

    if sum_sq_diff_x == 0.0 || sum_sq_diff_y == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    Ok(CellValue::Float(
        sum_prod_diff / (sum_sq_diff_x * sum_sq_diff_y).sqrt(),
    ))
}

pub(crate) async fn eval_pearson(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    // PEARSON is mathematically equivalent to CORREL
    eval_correl(ctx, current_sheet, args).await
}

pub(crate) async fn eval_rsq(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    match eval_pearson(ctx, current_sheet, args).await? {
        CellValue::Float(r) => Ok(CellValue::Float(r * r)),
        other => Ok(other),
    }
}

pub(crate) async fn eval_steyx(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("STEYX expects 2 arguments".to_string()));
    }

    let (ys, xs) = collect_aligned_numeric_pairs(ctx, current_sheet, &args[0], &args[1]).await?;
    let n = ys.len() as f64;
    if n < 3.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let mut sum_sq_diff_x = 0.0;
    let mut sum_sq_diff_y = 0.0;
    let mut sum_prod_diff = 0.0;

    for i in 0..ys.len() {
        let dx = xs[i] - mean_x;
        let dy = ys[i] - mean_y;
        sum_sq_diff_x += dx * dx;
        sum_sq_diff_y += dy * dy;
        sum_prod_diff += dx * dy;
    }

    if sum_sq_diff_x == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let term = sum_sq_diff_y - (sum_prod_diff * sum_prod_diff) / sum_sq_diff_x;
    // Handle potential precision issues leading to small negative values
    let term = term.max(0.0);

    Ok(CellValue::Float((term / (n - 2.0)).sqrt()))
}

pub(crate) async fn eval_slope(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error("SLOPE expects 2 arguments".to_string()));
    }

    let (ys, xs) = collect_aligned_numeric_pairs(ctx, current_sheet, &args[0], &args[1]).await?;
    let n = ys.len() as f64;
    if n < 1.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let mut sum_sq_diff_x = 0.0;
    let mut sum_prod_diff = 0.0;

    for i in 0..ys.len() {
        let dx = xs[i] - mean_x;
        let dy = ys[i] - mean_y;
        sum_sq_diff_x += dx * dx;
        sum_prod_diff += dx * dy;
    }

    if sum_sq_diff_x == 0.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    Ok(CellValue::Float(sum_prod_diff / sum_sq_diff_x))
}

pub(crate) async fn eval_intercept(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "INTERCEPT expects 2 arguments".to_string(),
        ));
    }

    let (ys, xs) = collect_aligned_numeric_pairs(ctx, current_sheet, &args[0], &args[1]).await?;
    let n = ys.len() as f64;
    if n < 1.0 {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    match eval_slope(ctx, current_sheet, args).await? {
        CellValue::Float(slope) => Ok(CellValue::Float(mean_y - slope * mean_x)),
        other => Ok(other),
    }
}

async fn eval_covariance(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    sample: bool,
) -> Result<CellValue> {
    if args.len() != 2 {
        let name = if sample {
            "COVARIANCE.S"
        } else {
            "COVARIANCE.P"
        };
        return Ok(CellValue::Error(format!("{} expects 2 arguments", name)));
    }

    let (xs, ys) = collect_aligned_numeric_pairs(ctx, current_sheet, &args[0], &args[1]).await?;
    let n = xs.len();
    if n == 0 || (sample && n < 2) {
        return Ok(CellValue::Error("#DIV/0!".to_string()));
    }

    let mean_x = xs.iter().sum::<f64>() / n as f64;
    let mean_y = ys.iter().sum::<f64>() / n as f64;

    let sum_prod_diff: f64 = xs
        .iter()
        .zip(ys.iter())
        .map(|(&x, &y)| (x - mean_x) * (y - mean_y))
        .sum();

    let divisor = if sample { (n - 1) as f64 } else { n as f64 };
    Ok(CellValue::Float(sum_prod_diff / divisor))
}

async fn collect_aligned_numeric_pairs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    arg1: &Expr,
    arg2: &Expr,
) -> Result<(Vec<f64>, Vec<f64>)> {
    use crate::sheet::eval::engine::flatten_range_expr;
    let range1 = flatten_range_expr(ctx, current_sheet, arg1).await?;
    let range2 = flatten_range_expr(ctx, current_sheet, arg2).await?;

    if range1.values.len() != range2.values.len() {
        // In Excel, this actually returns #N/A if they don't match, or sometimes #VALUE!
        // But usually they must be the same size.
        return Err(Box::new(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "Arrays must have the same size",
        )));
    }

    let mut xs = Vec::new();
    let mut ys = Vec::new();

    for (v1, v2) in range1.values.iter().zip(range2.values.iter()) {
        if let (Some(n1), Some(n2)) = (to_number(v1), to_number(v2)) {
            xs.push(n1);
            ys.push(n2);
        }
    }

    Ok((xs, ys))
}
