use super::{
    EvalCtx, FlatRange,
    criteria::{matches_criteria, parse_criteria},
    evaluate_expression, flatten_range_expr, is_blank, to_number, to_text,
};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(crate) async fn eval_dget(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "DGET expects 3 arguments (database, field, criteria)".to_string(),
        ));
    }

    let database = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    let rows = database.rows;
    let cols = database.cols;

    if rows < 2 || cols == 0 {
        return Ok(CellValue::Error(
            "DGET database must include header row and at least one record".to_string(),
        ));
    }

    let header_values = &database.values[..cols];
    let header_texts: Vec<String> = header_values.iter().map(to_text).collect();

    let field_arg = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let field_index = match field_to_index(&field_arg, &header_texts) {
        Some(idx) => idx,
        None => {
            return Ok(CellValue::Error(
                "DGET field must be a valid column header or 1-based index".to_string(),
            ));
        },
    };

    let criteria_range = flatten_range_expr(ctx, current_sheet, &args[2]).await?;
    if criteria_range.rows < 2 || criteria_range.cols == 0 {
        return Ok(CellValue::Error(
            "DGET criteria must include header row and at least one criteria row".to_string(),
        ));
    }

    let criteria_columns = match build_criteria_columns(&criteria_range, &header_texts) {
        Ok(cols) => cols,
        Err(err) => return Ok(err),
    };

    let matches = matching_records(&database, &criteria_range, &criteria_columns);
    match matches.len() {
        0 => Ok(CellValue::Error(
            "DGET found no rows matching criteria".to_string(),
        )),
        1 => Ok(matches[0][field_index].clone()),
        _ => Ok(CellValue::Error(
            "DGET found multiple rows matching criteria".to_string(),
        )),
    }
}

pub(crate) async fn eval_dmax(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DMAX", DatabaseStat::Max).await
}

pub(crate) async fn eval_dmin(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DMIN", DatabaseStat::Min).await
}

pub(crate) async fn eval_dcount(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(
        ctx,
        current_sheet,
        args,
        "DCOUNT",
        DatabaseStat::CountNumeric,
    )
    .await
}

pub(crate) async fn eval_dcounta(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DCOUNTA", DatabaseStat::CountAll).await
}

pub(crate) async fn eval_dsum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DSUM", DatabaseStat::Sum).await
}

pub(crate) async fn eval_dproduct(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DPRODUCT", DatabaseStat::Product).await
}

pub(crate) async fn eval_daverage(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DAVERAGE", DatabaseStat::Average).await
}

pub(crate) async fn eval_dstdev(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DSTDEV", DatabaseStat::StdSample).await
}

pub(crate) async fn eval_dstdevp(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(
        ctx,
        current_sheet,
        args,
        "DSTDEVP",
        DatabaseStat::StdPopulation,
    )
    .await
}

pub(crate) async fn eval_dvar(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(ctx, current_sheet, args, "DVAR", DatabaseStat::VarSample).await
}

pub(crate) async fn eval_dvarp(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    eval_database_stat(
        ctx,
        current_sheet,
        args,
        "DVARP",
        DatabaseStat::VarPopulation,
    )
    .await
}

enum DatabaseStat {
    Max,
    Min,
    CountNumeric,
    CountAll,
    Sum,
    Product,
    Average,
    StdSample,
    StdPopulation,
    VarSample,
    VarPopulation,
}

async fn eval_database_stat(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
    name: &str,
    mode: DatabaseStat,
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(format!(
            "{name} expects 3 arguments (database, field, criteria)"
        )));
    }

    let database = flatten_range_expr(ctx, current_sheet, &args[0]).await?;
    if database.rows < 2 || database.cols == 0 {
        return Ok(CellValue::Error(format!(
            "{name} database must include header row and at least one record"
        )));
    }

    let header_values = &database.values[..database.cols];
    let header_texts: Vec<String> = header_values.iter().map(to_text).collect();

    let field_arg = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let field_index = match field_to_index(&field_arg, &header_texts) {
        Some(idx) => idx,
        None => {
            return Ok(CellValue::Error(format!(
                "{name} field must be a valid column header or 1-based index"
            )));
        },
    };

    let criteria_range = flatten_range_expr(ctx, current_sheet, &args[2]).await?;
    if criteria_range.rows < 2 || criteria_range.cols == 0 {
        return Ok(CellValue::Error(format!(
            "{name} criteria must include header row and at least one criteria row"
        )));
    }

    let criteria_columns = match build_criteria_columns(&criteria_range, &header_texts) {
        Ok(cols) => cols,
        Err(err) => return Ok(err),
    };

    let matches = matching_records(&database, &criteria_range, &criteria_columns);
    if matches.is_empty() {
        return Ok(CellValue::Error(format!(
            "{name} found no rows matching criteria"
        )));
    }

    let mut numbers = Vec::new();
    let mut best_numeric: Option<f64> = None;
    let mut best_value: Option<CellValue> = None;
    let mut sum_acc = 0.0f64;
    let mut sum_found = false;
    let mut product_acc = 1.0f64;
    let mut product_found = false;
    let mut count_numeric = 0i64;
    let mut count_all = 0i64;
    let mut average_sum = 0.0f64;
    let mut average_count = 0i64;

    for row in matches {
        let value = &row[field_index];
        let numeric = to_number(value);

        match mode {
            DatabaseStat::Max => {
                if let Some(n) = numeric
                    && best_numeric.is_none_or(|current| n > current)
                {
                    best_numeric = Some(n);
                    best_value = Some(value.clone());
                }
            },
            DatabaseStat::Min => {
                if let Some(n) = numeric
                    && best_numeric.is_none_or(|current| n < current)
                {
                    best_numeric = Some(n);
                    best_value = Some(value.clone());
                }
            },
            DatabaseStat::CountNumeric => {
                if numeric.is_some() {
                    count_numeric += 1;
                }
            },
            DatabaseStat::CountAll => {
                if !is_blank(value) {
                    count_all += 1;
                }
            },
            DatabaseStat::Sum => {
                if let Some(n) = numeric {
                    sum_acc += n;
                    sum_found = true;
                }
            },
            DatabaseStat::Average => {
                if let Some(n) = numeric {
                    average_sum += n;
                    average_count += 1;
                }
            },
            DatabaseStat::Product => {
                if let Some(n) = numeric {
                    product_acc *= n;
                    product_found = true;
                }
            },
            DatabaseStat::StdSample
            | DatabaseStat::StdPopulation
            | DatabaseStat::VarSample
            | DatabaseStat::VarPopulation => {
                if let Some(n) = numeric {
                    numbers.push(n);
                }
            },
        }
    }

    match mode {
        DatabaseStat::Max => match best_value {
            Some(v) => Ok(v),
            None => Ok(CellValue::Error(format!(
                "{name} found no numeric values in the specified field"
            ))),
        },
        DatabaseStat::Min => match best_value {
            Some(v) => Ok(v),
            None => Ok(CellValue::Error(format!(
                "{name} found no numeric values in the specified field"
            ))),
        },
        DatabaseStat::CountNumeric => Ok(CellValue::Int(count_numeric)),
        DatabaseStat::CountAll => Ok(CellValue::Int(count_all)),
        DatabaseStat::Sum => {
            if !sum_found {
                return Ok(CellValue::Float(0.0));
            }
            Ok(CellValue::Float(sum_acc))
        },
        DatabaseStat::Average => {
            if average_count == 0 {
                return Ok(CellValue::Error(format!(
                    "{name} found no numeric values in the specified field"
                )));
            }
            Ok(CellValue::Float(average_sum / average_count as f64))
        },
        DatabaseStat::Product => {
            if !product_found {
                return Ok(CellValue::Float(0.0));
            }
            Ok(CellValue::Float(product_acc))
        },
        DatabaseStat::StdSample => {
            if numbers.len() < 2 {
                return Ok(CellValue::Error(format!(
                    "{name} requires at least two numeric records"
                )));
            }
            let mean = numbers.iter().sum::<f64>() / numbers.len() as f64;
            let variance = numbers.iter().map(|v| (v - mean).powi(2)).sum::<f64>()
                / (numbers.len() as f64 - 1.0);
            Ok(CellValue::Float(variance.sqrt()))
        },
        DatabaseStat::StdPopulation => {
            if numbers.is_empty() {
                return Ok(CellValue::Error(format!(
                    "{name} requires at least one numeric record"
                )));
            }
            let mean = numbers.iter().sum::<f64>() / numbers.len() as f64;
            let variance =
                numbers.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / numbers.len() as f64;
            Ok(CellValue::Float(variance.sqrt()))
        },
        DatabaseStat::VarSample => {
            if numbers.len() < 2 {
                return Ok(CellValue::Error(format!(
                    "{name} requires at least two numeric records"
                )));
            }
            let mean = numbers.iter().sum::<f64>() / numbers.len() as f64;
            let variance = numbers.iter().map(|v| (v - mean).powi(2)).sum::<f64>()
                / (numbers.len() as f64 - 1.0);
            Ok(CellValue::Float(variance))
        },
        DatabaseStat::VarPopulation => {
            if numbers.is_empty() {
                return Ok(CellValue::Error(format!(
                    "{name} requires at least one numeric record"
                )));
            }
            let mean = numbers.iter().sum::<f64>() / numbers.len() as f64;
            let variance =
                numbers.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / numbers.len() as f64;
            Ok(CellValue::Float(variance))
        },
    }
}

fn field_to_index(field: &CellValue, headers: &[String]) -> Option<usize> {
    match field {
        CellValue::Int(i) => {
            if *i <= 0 {
                return None;
            }
            let idx = (*i - 1) as usize;
            (idx < headers.len()).then_some(idx)
        },
        _ => {
            let name = to_text(field);
            if name.is_empty() {
                None
            } else {
                headers.iter().position(|h| h == &name)
            }
        },
    }
}

fn build_criteria_columns(
    criteria: &FlatRange,
    headers: &[String],
) -> std::result::Result<Vec<Option<usize>>, CellValue> {
    let mut columns = Vec::with_capacity(criteria.cols);
    for c in 0..criteria.cols {
        let label = to_text(&criteria.values[c]);
        if label.is_empty() {
            columns.push(None);
            continue;
        }
        match headers.iter().position(|h| h == &label) {
            Some(idx) => columns.push(Some(idx)),
            None => {
                return Err(CellValue::Error(format!(
                    "DGET criteria column '{}' not found in database headers",
                    label
                )));
            },
        }
    }
    Ok(columns)
}

fn record_matches(
    record: &[CellValue],
    criteria_range: &FlatRange,
    column_map: &[Option<usize>],
) -> bool {
    let criteria_cols = criteria_range.cols;
    for r in 1..criteria_range.rows {
        let mut row_ok = true;
        let mut has_condition = false;
        for (c, column_index_opt) in column_map.iter().enumerate() {
            let crit_value = &criteria_range.values[r * criteria_cols + c];
            if is_blank(crit_value) {
                continue;
            }

            let column_index = match column_index_opt {
                Some(idx) => *idx,
                None => {
                    row_ok = false;
                    break;
                },
            };

            has_condition = true;
            let crit_text = to_text(crit_value);
            let criteria = match parse_criteria(&crit_text) {
                Some(c) => c,
                None => {
                    row_ok = false;
                    break;
                },
            };

            if column_index >= record.len() || !matches_criteria(&record[column_index], &criteria) {
                row_ok = false;
                break;
            }
        }

        if row_ok && has_condition {
            return true;
        }
    }

    false
}

fn matching_records<'a>(
    database: &'a FlatRange,
    criteria_range: &FlatRange,
    column_map: &[Option<usize>],
) -> Vec<&'a [CellValue]> {
    let mut rows = Vec::new();
    let cols = database.cols;
    for r in 1..database.rows {
        let row_slice = &database.values[r * cols..(r + 1) * cols];
        if record_matches(row_slice, criteria_range, column_map) {
            rows.push(row_slice);
        }
    }
    rows
}
