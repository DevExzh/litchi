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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::engine::test_helpers::TestEngine;
    use crate::sheet::eval::parser::{Expr, RangeRef};

    fn create_database_range() -> Expr {
        // Create a database with headers: Name, Age, Score
        // Records: Alice(25, 85), Bob(30, 90), Carol(25, 78), Dave(35, 92)
        Expr::Range(RangeRef {
            sheet: "Sheet1".to_string(),
            start_col: 0,
            start_row: 0,
            end_col: 2,
            end_row: 4,
        })
    }

    fn setup_database(engine: &TestEngine) {
        // Header row
        engine.set_cell("Sheet1", 0, 0, CellValue::String("Name".to_string()));
        engine.set_cell("Sheet1", 0, 1, CellValue::String("Age".to_string()));
        engine.set_cell("Sheet1", 0, 2, CellValue::String("Score".to_string()));
        // Data rows
        engine.set_cell("Sheet1", 1, 0, CellValue::String("Alice".to_string()));
        engine.set_cell("Sheet1", 1, 1, CellValue::Int(25));
        engine.set_cell("Sheet1", 1, 2, CellValue::Int(85));
        engine.set_cell("Sheet1", 2, 0, CellValue::String("Bob".to_string()));
        engine.set_cell("Sheet1", 2, 1, CellValue::Int(30));
        engine.set_cell("Sheet1", 2, 2, CellValue::Int(90));
        engine.set_cell("Sheet1", 3, 0, CellValue::String("Carol".to_string()));
        engine.set_cell("Sheet1", 3, 1, CellValue::Int(25));
        engine.set_cell("Sheet1", 3, 2, CellValue::Int(78));
        engine.set_cell("Sheet1", 4, 0, CellValue::String("Dave".to_string()));
        engine.set_cell("Sheet1", 4, 1, CellValue::Int(35));
        engine.set_cell("Sheet1", 4, 2, CellValue::Int(92));
    }

    fn create_criteria_range(criteria_col: u32, criteria_row: u32) -> Expr {
        Expr::Range(RangeRef {
            sheet: "Sheet1".to_string(),
            start_col: criteria_col,
            start_row: criteria_row,
            end_col: criteria_col + 2,
            end_row: criteria_row + 1,
        })
    }

    fn setup_age_criteria(
        engine: &TestEngine,
        criteria_col: u32,
        criteria_row: u32,
        age_value: i64,
    ) {
        // Criteria header
        engine.set_cell(
            "Sheet1",
            criteria_row,
            criteria_col,
            CellValue::String("Age".to_string()),
        );
        // Criteria value
        engine.set_cell(
            "Sheet1",
            criteria_row + 1,
            criteria_col,
            CellValue::Int(age_value),
        );
    }

    #[tokio::test]
    async fn test_dget_single_match() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 30); // Age = 30

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(90)); // Bob's score
    }

    #[tokio::test]
    async fn test_dget_no_match() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 99); // Age = 99 (no match)

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("found no rows")),
            _ => panic!("Expected Error result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dget_multiple_matches() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 25); // Age = 25 (Alice and Carol)

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("multiple rows")),
            _ => panic!("Expected Error result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dget_field_by_index() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 30);

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::Int(3)); // 3rd column (Score)
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(90));
    }

    #[tokio::test]
    async fn test_dsum() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 25); // Age = 25 (Alice=85, Carol=78)

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dsum(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Float(163.0)); // 85 + 78
    }

    #[tokio::test]
    async fn test_daverage() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 25); // Age = 25

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_daverage(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Float(81.5)); // (85 + 78) / 2
    }

    #[tokio::test]
    async fn test_dcount() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 25);

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dcount(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(2)); // 2 numeric scores
    }

    #[tokio::test]
    async fn test_dcounta() {
        let engine = TestEngine::new();
        setup_database(&engine);
        // Add criteria to match all rows (using Score > 0)
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Name".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dcounta(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(4)); // 4 names match Score > 0
    }

    #[tokio::test]
    async fn test_dmax() {
        let engine = TestEngine::new();
        setup_database(&engine);
        // Match all by using criteria with blank value
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dmax(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(92)); // Dave's score
    }

    #[tokio::test]
    async fn test_dmin() {
        let engine = TestEngine::new();
        setup_database(&engine);
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dmin(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(78)); // Carol's score
    }

    #[tokio::test]
    async fn test_dproduct() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 25); // Age = 25 (scores 85, 78)

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dproduct(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Float(6630.0)); // 85 * 78
    }

    #[tokio::test]
    async fn test_dstddev() {
        let engine = TestEngine::new();
        setup_database(&engine);
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dstdev(ctx, "Sheet1", &args).await.unwrap();
        // Sample standard deviation of [85, 90, 78, 92]
        // mean = 86.25, variance = 38.916667, stddev = 6.238322
        match result {
            CellValue::Float(v) => {
                let expected = 6.23832242407;
                assert!(
                    (v - expected).abs() < 0.001,
                    "Expected ~{}, got {}",
                    expected,
                    v
                );
            },
            _ => panic!("Expected Float result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dstddevp() {
        let engine = TestEngine::new();
        setup_database(&engine);
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dstdevp(ctx, "Sheet1", &args).await.unwrap();
        // Population standard deviation of [85, 90, 78, 92]
        // mean = 86.25, variance = 29.1875, stddev = 5.402546
        match result {
            CellValue::Float(v) => {
                let expected = 5.40254569624;
                assert!(
                    (v - expected).abs() < 0.001,
                    "Expected ~{}, got {}",
                    expected,
                    v
                );
            },
            _ => panic!("Expected Float result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dvar() {
        let engine = TestEngine::new();
        setup_database(&engine);
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dvar(ctx, "Sheet1", &args).await.unwrap();
        // Sample variance of [85, 90, 78, 92] = 38.916667
        match result {
            CellValue::Float(v) => {
                let expected = 38.9166666667;
                assert!(
                    (v - expected).abs() < 0.001,
                    "Expected ~{}, got {}",
                    expected,
                    v
                );
            },
            _ => panic!("Expected Float result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dvarp() {
        let engine = TestEngine::new();
        setup_database(&engine);
        engine.set_cell("Sheet1", 0, 10, CellValue::String("Score".to_string()));
        engine.set_cell("Sheet1", 1, 10, CellValue::String(">0".to_string()));

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("Score".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dvarp(ctx, "Sheet1", &args).await.unwrap();
        // Population variance of [85, 90, 78, 92] = 29.1875
        match result {
            CellValue::Float(v) => {
                let expected = 29.1875;
                assert!(
                    (v - expected).abs() < 0.001,
                    "Expected ~{}, got {}",
                    expected,
                    v
                );
            },
            _ => panic!("Expected Float result, got {:?}", result),
        }
    }

    #[tokio::test]
    async fn test_dget_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 3 arguments")),
            _ => panic!("Expected Error result"),
        }
    }

    #[tokio::test]
    async fn test_dget_invalid_field() {
        let engine = TestEngine::new();
        setup_database(&engine);
        setup_age_criteria(&engine, 10, 0, 30);

        let ctx = engine.ctx();
        let database = create_database_range();
        let field = Expr::Literal(CellValue::String("InvalidField".to_string()));
        let criteria = create_criteria_range(10, 0);

        let args = vec![database, field, criteria];
        let result = eval_dget(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("valid column header")),
            _ => panic!("Expected Error result, got {:?}", result),
        }
    }
}
