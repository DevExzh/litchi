use std::cmp::Ordering;
use std::result::Result as StdResult;

use crate::sheet::eval::engine::{EvalCtx, flatten_range_expr};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    EPS, collect_numeric_values, collect_numeric_values_unsorted, number_arg, to_positive_index,
};

pub(crate) async fn eval_large(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    extremum_function(ctx, current_sheet, args, "LARGE", ExtremumKind::Large).await
}

pub(crate) async fn eval_small(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    extremum_function(ctx, current_sheet, args, "SMALL", ExtremumKind::Small).await
}

pub(crate) async fn eval_rank_eq(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    rank_function(ctx, current_sheet, args, "RANK.EQ", RankType::Equal).await
}

pub(crate) async fn eval_rank_avg(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    rank_function(ctx, current_sheet, args, "RANK.AVG", RankType::Average).await
}

pub(crate) async fn eval_rank(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    rank_function(ctx, current_sheet, args, "RANK", RankType::Equal).await
}

pub(crate) async fn eval_percentile(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentile_function(
        ctx,
        current_sheet,
        args,
        "PERCENTILE",
        PercentileMode::Inclusive,
    )
    .await
}

pub(crate) async fn eval_percentile_inc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentile_function(
        ctx,
        current_sheet,
        args,
        "PERCENTILE.INC",
        PercentileMode::Inclusive,
    )
    .await
}

pub(crate) async fn eval_percentile_exc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentile_function(
        ctx,
        current_sheet,
        args,
        "PERCENTILE.EXC",
        PercentileMode::Exclusive,
    )
    .await
}

pub(crate) async fn eval_quartile(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    quartile_function(
        ctx,
        current_sheet,
        args,
        "QUARTILE",
        QuartileMode::Inclusive,
    )
    .await
}

pub(crate) async fn eval_quartile_inc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    quartile_function(
        ctx,
        current_sheet,
        args,
        "QUARTILE.INC",
        QuartileMode::Inclusive,
    )
    .await
}

pub(crate) async fn eval_quartile_exc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    quartile_function(
        ctx,
        current_sheet,
        args,
        "QUARTILE.EXC",
        QuartileMode::Exclusive,
    )
    .await
}

pub(crate) async fn eval_percentrank(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentrank_function(ctx, current_sheet, args, "PERCENTRANK", RankMode::Inclusive).await
}

pub(crate) async fn eval_percentrank_inc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentrank_function(
        ctx,
        current_sheet,
        args,
        "PERCENTRANK.INC",
        RankMode::Inclusive,
    )
    .await
}

pub(crate) async fn eval_percentrank_exc(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    percentrank_function(
        ctx,
        current_sheet,
        args,
        "PERCENTRANK.EXC",
        RankMode::Exclusive,
    )
    .await
}

#[derive(Clone, Copy)]
enum PercentileMode {
    Inclusive,
    Exclusive,
}

#[derive(Clone, Copy)]
enum QuartileMode {
    Inclusive,
    Exclusive,
}

#[derive(Clone, Copy)]
enum RankMode {
    Inclusive,
    Exclusive,
}

#[derive(Clone, Copy)]
enum ExtremumKind {
    Large,
    Small,
}

#[derive(Clone, Copy)]
enum RankType {
    Equal,
    Average,
}

async fn extremum_function(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    kind: ExtremumKind,
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 2 arguments (array, k)"
        )));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let numbers = match collect_numeric_values(array.values, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let k_value = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error(format!("{func_name} k must be numeric"))),
    };

    let k = match to_positive_index(k_value, func_name, "k") {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    if k > numbers.len() {
        return Ok(CellValue::Error(format!(
            "{func_name} k cannot exceed the number of numeric values in the array"
        )));
    }

    let idx = match kind {
        ExtremumKind::Large => numbers.len() - k,
        ExtremumKind::Small => k - 1,
    };

    Ok(CellValue::Float(numbers[idx]))
}

async fn rank_function(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    rank_type: RankType,
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 2 or 3 arguments (number, ref, [order])"
        )));
    }

    let number = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(format!(
                "{func_name} number must be numeric"
            )));
        },
    };

    if !number.is_finite() {
        return Ok(CellValue::Error(format!(
            "{func_name} number must be finite"
        )));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[1]).await?;
    let values = match collect_numeric_values_unsorted(array.values, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };
    let ascending = if args.len() == 3 {
        match number_arg(ctx, current_sheet, &args[2]).await? {
            Some(v) => v != 0.0,
            None => {
                return Ok(CellValue::Error(format!(
                    "{func_name} order must be numeric"
                )));
            },
        }
    } else {
        false
    };

    if std::env::var("RANK_DEBUG").is_ok() {
        eprintln!(
            "rank_function input: func={func_name}, number={number}, ascending={ascending}, values={values:?}"
        );
    }

    match rank_type {
        RankType::Equal => rank_equal_value(number, &values, ascending, func_name),
        RankType::Average => rank_average_value(number, &values, ascending, func_name),
    }
}

async fn percentile_function(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    mode: PercentileMode,
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 2 arguments (array, k)"
        )));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let numbers = match collect_numeric_values(array.values, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let k = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => return Ok(CellValue::Error(format!("{func_name} k must be numeric"))),
    };

    let value = match percentile_value(&numbers, k, mode, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    Ok(CellValue::Float(value))
}

async fn quartile_function(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    mode: QuartileMode,
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 2 arguments (array, quart)"
        )));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let numbers = match collect_numeric_values(array.values, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let quart_value = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(format!(
                "{func_name} quart argument must be numeric"
            )));
        },
    };

    let quart_int = match to_int_quart(quart_value, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let (percent_mode, k) = match mode {
        QuartileMode::Inclusive => {
            if !(0..=4).contains(&quart_int) {
                return Ok(CellValue::Error(format!(
                    "{func_name} quart must be between 0 and 4"
                )));
            }
            (PercentileMode::Inclusive, quart_int as f64 / 4.0)
        },
        QuartileMode::Exclusive => {
            if !(1..=3).contains(&quart_int) {
                return Ok(CellValue::Error(format!(
                    "{func_name} quart must be 1, 2, or 3"
                )));
            }
            (PercentileMode::Exclusive, quart_int as f64 / 4.0)
        },
    };

    let value = match percentile_value(&numbers, k, percent_mode, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    Ok(CellValue::Float(value))
}

async fn percentrank_function(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    func_name: &str,
    mode: RankMode,
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(format!(
            "{func_name} expects 2 or 3 arguments (array, x, [significance])"
        )));
    }

    let array = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let numbers = match collect_numeric_values(array.values, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let x = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(v) => v,
        None => {
            return Ok(CellValue::Error(format!(
                "{func_name} requires x to be numeric"
            )));
        },
    };

    let significance = if args.len() == 3 {
        match number_arg(ctx, current_sheet, &args[2]).await? {
            Some(v) => match parse_significance(v, func_name) {
                Ok(s) => s,
                Err(err) => return Ok(err),
            },
            None => {
                return Ok(CellValue::Error(format!(
                    "{func_name} significance must be numeric"
                )));
            },
        }
    } else {
        3
    };

    let rank = match percentrank_value(&numbers, x, mode, func_name) {
        Ok(v) => v,
        Err(err) => return Ok(err),
    };

    let rounded = round_to_significance(rank, significance);
    Ok(CellValue::Float(rounded))
}

fn rank_equal_value(
    number: f64,
    values: &[f64],
    ascending: bool,
    func_name: &str,
) -> Result<CellValue> {
    let positions = match rank_positions(number, values, ascending, func_name) {
        Ok(p) => p,
        Err(err) => return Ok(err),
    };
    Ok(CellValue::Float(positions.first as f64))
}

fn rank_average_value(
    number: f64,
    values: &[f64],
    ascending: bool,
    func_name: &str,
) -> Result<CellValue> {
    let positions = match rank_positions(number, values, ascending, func_name) {
        Ok(p) => p,
        Err(err) => return Ok(err),
    };
    let avg = (positions.first + positions.last) as f64 / 2.0;
    Ok(CellValue::Float(avg))
}

struct RankPositions {
    first: usize,
    last: usize,
}

fn rank_positions(
    number: f64,
    values: &[f64],
    ascending: bool,
    func_name: &str,
) -> StdResult<RankPositions, CellValue> {
    if std::env::var("RANK_DEBUG").is_ok() {
        eprintln!(
            "rank_positions input: func={func_name}, ascending={ascending}, number={number}, values={values:?}"
        );
    }
    let mut indexed: Vec<(usize, f64)> = values
        .iter()
        .enumerate()
        .filter(|(_, v)| v.is_finite())
        .map(|(i, &v)| (i, v))
        .collect();

    if indexed.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} reference contains no numeric values"
        )));
    }

    indexed.sort_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(Ordering::Equal));
    if !ascending {
        indexed.reverse();
    }
    if ascending && std::env::var("RANK_DEBUG").is_ok() {
        let ordered: Vec<f64> = indexed.iter().map(|(_, v)| *v).collect();
        eprintln!(
            "rank_positions ordered: func={func_name}, ascending={ascending}, number={number}, ordered={ordered:?}"
        );
    }

    let mut first: Option<usize> = None;
    let mut last: Option<usize> = None;

    for (rank, (_, value)) in indexed.iter().enumerate() {
        if approx_equal(*value, number) {
            let display_rank = rank + 1;
            if first.is_none() {
                first = Some(display_rank);
            }
            last = Some(display_rank);
        }
    }

    match (first, last) {
        (Some(f), Some(l)) => {
            if ascending && std::env::var("RANK_DEBUG").is_ok() {
                eprintln!(
                    "rank_positions result: func={func_name}, ascending={ascending}, number={number}, first={f}, last={l}"
                );
            }
            Ok(RankPositions { first: f, last: l })
        },
        _ => Err(CellValue::Error(format!(
            "{func_name} number must exist in the reference"
        ))),
    }
}

fn approx_equal(a: f64, b: f64) -> bool {
    (a - b).abs() < EPS
}

fn percentile_value(
    values: &[f64],
    k: f64,
    mode: PercentileMode,
    func_name: &str,
) -> StdResult<f64, CellValue> {
    if k.is_nan() {
        return Err(CellValue::Error(format!(
            "{func_name} k must be between 0 and 1"
        )));
    }

    match mode {
        PercentileMode::Inclusive => {
            if !(0.0..=1.0).contains(&k) {
                return Err(CellValue::Error(format!(
                    "{func_name} k must be between 0 and 1"
                )));
            }
            if values.len() == 1 {
                return Ok(values[0]);
            }
            let pos = k * (values.len() as f64 - 1.0);
            Ok(interpolate_sorted(values, pos))
        },
        PercentileMode::Exclusive => {
            if k <= 0.0 || k >= 1.0 {
                return Err(CellValue::Error(format!(
                    "{func_name} k must be between 0 and 1 (exclusive)"
                )));
            }
            if values.len() < 2 {
                return Err(CellValue::Error(format!(
                    "{func_name} requires at least two numeric values"
                )));
            }
            let pos = k * (values.len() as f64 + 1.0);
            if pos <= 1.0 || pos >= values.len() as f64 {
                return Err(CellValue::Error(format!(
                    "{func_name} k results in an invalid rank"
                )));
            }
            let zero_based = pos - 1.0;
            Ok(interpolate_sorted(values, zero_based))
        },
    }
}

fn interpolate_sorted(values: &[f64], pos: f64) -> f64 {
    let lower = pos.floor();
    let upper = pos.ceil();
    let lower_idx = lower.clamp(0.0, (values.len() - 1) as f64) as usize;
    let upper_idx = upper.clamp(0.0, (values.len() - 1) as f64) as usize;

    if lower_idx == upper_idx {
        return values[lower_idx];
    }

    let fraction = pos - lower;
    let lower_value = values[lower_idx];
    let upper_value = values[upper_idx];
    lower_value + fraction * (upper_value - lower_value)
}

fn to_int_quart(value: f64, func_name: &str) -> StdResult<i32, CellValue> {
    let rounded = value.round();
    if (value - rounded).abs() > EPS {
        return Err(CellValue::Error(format!(
            "{func_name} quart argument must be an integer"
        )));
    }
    if rounded < i32::MIN as f64 || rounded > i32::MAX as f64 {
        return Err(CellValue::Error(format!(
            "{func_name} quart argument is out of range"
        )));
    }
    Ok(rounded as i32)
}

fn parse_significance(value: f64, func_name: &str) -> StdResult<u32, CellValue> {
    if !value.is_finite() {
        return Err(CellValue::Error(format!(
            "{func_name} significance must be a finite positive integer"
        )));
    }
    let rounded = value.round();
    if (value - rounded).abs() > EPS || rounded <= 0.0 {
        return Err(CellValue::Error(format!(
            "{func_name} significance must be a positive integer"
        )));
    }
    if rounded > 15.0 {
        return Err(CellValue::Error(format!(
            "{func_name} significance cannot exceed 15"
        )));
    }
    Ok(rounded as u32)
}

fn percentrank_value(
    values: &[f64],
    x: f64,
    mode: RankMode,
    func_name: &str,
) -> StdResult<f64, CellValue> {
    if !x.is_finite() {
        return Err(CellValue::Error(format!(
            "{func_name} x must be a finite numeric value"
        )));
    }
    if values.len() == 1 {
        if (x - values[0]).abs() < EPS {
            return Ok(0.0);
        }
        return Err(CellValue::Error(format!(
            "{func_name} cannot rank x outside the data range"
        )));
    }

    let min = values[0];
    let max = values[values.len() - 1];

    match mode {
        RankMode::Inclusive => {
            if x < min || x > max {
                return Err(CellValue::Error(format!(
                    "{func_name} x must lie within the data array"
                )));
            }
            let position = match inclusive_position(values, x) {
                Some(p) => p,
                None => {
                    return Err(CellValue::Error(format!(
                        "{func_name} could not determine rank position"
                    )));
                },
            };
            let denom = (values.len() - 1) as f64;
            if denom.abs() < EPS {
                return Ok(0.0);
            }
            Ok((position / denom).clamp(0.0, 1.0))
        },
        RankMode::Exclusive => {
            if x <= min || x >= max {
                return Err(CellValue::Error(format!(
                    "{func_name} x must be between the minimum and maximum values"
                )));
            }
            if values.len() < 2 {
                return Err(CellValue::Error(format!(
                    "{func_name} requires at least two numeric values"
                )));
            }
            let position = match inclusive_position(values, x) {
                Some(p) => p,
                None => {
                    return Err(CellValue::Error(format!(
                        "{func_name} could not determine rank position"
                    )));
                },
            };
            let denom = (values.len() + 1) as f64;
            let shifted = position + 1.0;
            if shifted <= 1.0 || shifted >= values.len() as f64 {
                return Err(CellValue::Error(format!(
                    "{func_name} x results in an invalid rank"
                )));
            }
            Ok((shifted / denom).clamp(0.0, 1.0))
        },
    }
}

fn inclusive_position(values: &[f64], x: f64) -> Option<f64> {
    match values.binary_search_by(|v| v.partial_cmp(&x).unwrap_or(Ordering::Equal)) {
        Ok(idx) => Some(idx as f64),
        Err(insert_pos) => {
            if insert_pos == 0 || insert_pos >= values.len() {
                return None;
            }
            let lower = values[insert_pos - 1];
            let upper = values[insert_pos];
            if (upper - lower).abs() < EPS {
                Some((insert_pos - 1) as f64)
            } else {
                Some((insert_pos - 1) as f64 + (x - lower) / (upper - lower))
            }
        },
    }
}

fn round_to_significance(value: f64, significance: u32) -> f64 {
    let factor = 10f64.powi(significance as i32);
    (value * factor).round() / factor
}
