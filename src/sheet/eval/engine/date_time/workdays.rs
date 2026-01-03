use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    WeekendConfig, collect_holiday_dates, date_to_excel_serial_1900, networkdays_core, number_arg,
    serial_to_excel_date_1900, workday_core,
};

pub(crate) async fn eval_workday(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WORKDAY expects 2 or 3 arguments (start_date, days, [holidays])".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY start_date is not numeric".to_string(),
            ));
        },
    };
    let days = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("WORKDAY days is not numeric".to_string()));
        },
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY start_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 3 {
        collect_holiday_dates(ctx, current_sheet, &args[2]).await?
    } else {
        Vec::new()
    };

    let weekend = WeekendConfig::default();
    let result = match workday_core(start_date, days.trunc() as i64, &holidays, &weekend) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(result);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_workday_intl(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "WORKDAY.INTL expects 2 to 4 arguments (start_date, days, [weekend], [holidays])"
                .to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL start_date is not numeric".to_string(),
            ));
        },
    };
    let days = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL days is not numeric".to_string(),
            ));
        },
    };

    let weekend = if args.len() >= 3 {
        let weekend_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match weekend_val {
            CellValue::String(s) => WeekendConfig::Pattern(s),
            _ => match to_number(&weekend_val) {
                Some(n) => WeekendConfig::Code(n.trunc() as i32),
                None => return Ok(CellValue::Error("#VALUE!".to_string())),
            },
        }
    } else {
        WeekendConfig::default()
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL start_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 4 {
        collect_holiday_dates(ctx, current_sheet, &args[3]).await?
    } else {
        Vec::new()
    };

    let result = match workday_core(start_date, days.trunc() as i64, &holidays, &weekend) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(result);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_networkdays(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "NETWORKDAYS expects 2 or 3 arguments (start_date, end_date, [holidays])".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS start_date is not numeric".to_string(),
            ));
        },
    };
    let end = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS end_date is not numeric".to_string(),
            ));
        },
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS start_date is not a valid date".to_string(),
            ));
        },
    };
    let end_date = match serial_to_excel_date_1900(end) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS end_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 3 {
        collect_holiday_dates(ctx, current_sheet, &args[2]).await?
    } else {
        Vec::new()
    };

    let weekend = WeekendConfig::default();
    let count = networkdays_core(start_date, end_date, &holidays, &weekend);
    Ok(CellValue::Int(count))
}

pub(crate) async fn eval_networkdays_intl(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "NETWORKDAYS.INTL expects 2 to 4 arguments (start_date, end_date, [weekend], [holidays])"
                .to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS.INTL start_date is not numeric".to_string(),
            ));
        },
    };
    let end = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS.INTL end_date is not numeric".to_string(),
            ));
        },
    };

    let weekend = if args.len() >= 3 {
        let weekend_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match weekend_val {
            CellValue::String(s) => WeekendConfig::Pattern(s),
            _ => match to_number(&weekend_val) {
                Some(n) => WeekendConfig::Code(n.trunc() as i32),
                None => return Ok(CellValue::Error("#VALUE!".to_string())),
            },
        }
    } else {
        WeekendConfig::default()
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS.INTL start_date is not a valid date".to_string(),
            ));
        },
    };
    let end_date = match serial_to_excel_date_1900(end) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS.INTL end_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 4 {
        collect_holiday_dates(ctx, current_sheet, &args[3]).await?
    } else {
        Vec::new()
    };

    let count = networkdays_core(start_date, end_date, &holidays, &weekend);
    Ok(CellValue::Int(count))
}
