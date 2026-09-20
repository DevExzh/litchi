//! Scalar façade and shared kernels for OpenFormula section 6.10 date/time.
//!
//! Resolver-backed sequence traversal is implemented by the value evaluator.
//! This module owns the function catalog, scalar coercion/error ordering, and
//! the pure date/time operations exported to that bridge.

pub(super) mod kernel;
mod parser;

use super::{
    EvaluationFailure, EvaluationResult, Evaluator, Node, ScalarError, TextValue, UnsupportedKind,
    WorkingValue,
};
use kernel::CivilDate;

pub(super) use kernel::{
    MAX_DATETIME_SERIAL, MIN_SERIAL, add_months, date_difference, date_to_serial, days_from_civil,
    days_in_month, easter_sunday, is_workday, iso_week, serial_to_date, serial_to_parts,
    weekday_sunday_zero, yearfrac,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Function {
    Date,
    DateDif,
    DateValue,
    Day,
    Days,
    Days360,
    EasterSunday,
    EDate,
    EOMonth,
    Hour,
    IsoWeeknum,
    Minute,
    Month,
    Networkdays,
    Now,
    Second,
    Time,
    TimeValue,
    Today,
    Weekday,
    Weeknum,
    Workday,
    Year,
    Yearfrac,
}

impl Function {
    pub(super) fn from_name(name: &str) -> Option<Self> {
        Some(match name.len() {
            3 if name.eq_ignore_ascii_case("DAY") => Self::Day,
            4 if name.eq_ignore_ascii_case("DAYS") => Self::Days,
            4 if name.eq_ignore_ascii_case("DATE") => Self::Date,
            4 if name.eq_ignore_ascii_case("HOUR") => Self::Hour,
            5 if name.eq_ignore_ascii_case("MONTH") => Self::Month,
            3 if name.eq_ignore_ascii_case("NOW") => Self::Now,
            4 if name.eq_ignore_ascii_case("TIME") => Self::Time,
            4 if name.eq_ignore_ascii_case("YEAR") => Self::Year,
            5 if name.eq_ignore_ascii_case("EDATE") => Self::EDate,
            6 if name.eq_ignore_ascii_case("MINUTE") => Self::Minute,
            6 if name.eq_ignore_ascii_case("SECOND") => Self::Second,
            5 if name.eq_ignore_ascii_case("TODAY") => Self::Today,
            7 if name.eq_ignore_ascii_case("DAYS360") => Self::Days360,
            7 if name.eq_ignore_ascii_case("WEEKDAY") => Self::Weekday,
            7 if name.eq_ignore_ascii_case("DATEDIF") => Self::DateDif,
            9 if name.eq_ignore_ascii_case("DATEVALUE") => Self::DateValue,
            7 if name.eq_ignore_ascii_case("EOMONTH") => Self::EOMonth,
            10 if name.eq_ignore_ascii_case("ISOWEEKNUM") => Self::IsoWeeknum,
            7 if name.eq_ignore_ascii_case("WEEKNUM") => Self::Weeknum,
            8 if name.eq_ignore_ascii_case("YEARFRAC") => Self::Yearfrac,
            7 if name.eq_ignore_ascii_case("WORKDAY") => Self::Workday,
            11 if name.eq_ignore_ascii_case("NETWORKDAYS") => Self::Networkdays,
            12 if name.eq_ignore_ascii_case("EASTERSUNDAY") => Self::EasterSunday,
            9 if name.eq_ignore_ascii_case("TIMEVALUE") => Self::TimeValue,
            _ => return None,
        })
    }

    pub(super) const fn valid_arity(self, count: usize) -> bool {
        match self {
            Self::Date | Self::DateDif | Self::Time => count == 3,
            Self::DateValue
            | Self::Day
            | Self::Hour
            | Self::IsoWeeknum
            | Self::Minute
            | Self::Month
            | Self::Second
            | Self::TimeValue
            | Self::Year => count == 1,
            Self::Days | Self::EDate | Self::EOMonth => count == 2,
            Self::Days360 | Self::Yearfrac => count >= 2 && count <= 3,
            Self::Weekday | Self::Weeknum => count >= 1 && count <= 2,
            Self::EasterSunday => count <= 1,
            Self::Now | Self::Today => count == 0,
            Self::Networkdays | Self::Workday => count >= 2 && count <= 4,
        }
    }

    pub(super) const fn sequence_argument(self, index: usize) -> bool {
        matches!(self, Self::Networkdays | Self::Workday) && (index == 2 || index == 3)
    }
}

pub(super) fn is_date_time_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

fn reverse_value_tail<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    count: usize,
) -> EvaluationResult<()> {
    evaluator.charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
    let start =
        evaluator
            .values
            .len()
            .checked_sub(count)
            .ok_or(EvaluationFailure::InvalidExpression(
                "date/time value stack underflow",
            ))?;
    evaluator.values[start..].reverse();
    Ok(())
}

struct Arguments<'a> {
    values: [Option<WorkingValue<'a>>; 4],
    formula_error: Option<ScalarError>,
}

fn collect_arguments<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
) -> EvaluationResult<Arguments<'a>> {
    let count = node.child_count();
    reverse_value_tail(evaluator, count)?;
    let mut values: [Option<WorkingValue<'a>>; 4] = std::array::from_fn(|_| None);
    let mut formula_error = None;
    for (index, slot) in values.iter_mut().enumerate().take(count) {
        let value = evaluator.pop_value()?;
        let missing = node.child(index).is_some_and(|child| child.is_missing());
        if !missing && formula_error.is_none() {
            if let WorkingValue::Error(error) = value {
                formula_error = Some(error);
            }
        }
        if !missing {
            *slot = Some(value);
        }
    }
    Ok(Arguments {
        values,
        formula_error,
    })
}

fn required<'a>(args: &mut Arguments<'a>, index: usize) -> Result<WorkingValue<'a>, ScalarError> {
    args.values
        .get_mut(index)
        .and_then(Option::take)
        .ok_or(ScalarError::Value)
}

fn optional<'a>(args: &mut Arguments<'a>, index: usize) -> Option<WorkingValue<'a>> {
    args.values.get_mut(index).and_then(Option::take)
}

fn push_result<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    result: Result<f64, ScalarError>,
) -> EvaluationResult<()> {
    match result {
        Ok(value) if value.is_finite() => evaluator.push_value(WorkingValue::Number(value)),
        Ok(_) => evaluator.push_value(WorkingValue::Error(ScalarError::Number)),
        Err(error) => evaluator.push_value(WorkingValue::Error(error)),
    }
}

fn number_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    super::to_number(value, evaluator)
}

fn integer_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<i64, ScalarError>> {
    let value = match number_argument(evaluator, value)? {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    if !value.is_finite()
        || !(-9_223_372_036_854_775_808.0..9_223_372_036_854_775_808.0).contains(&value)
    {
        return Ok(Err(ScalarError::Number));
    }
    Ok(Ok(value.trunc() as i64))
}

fn logical_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<bool, ScalarError>> {
    super::to_logical(value, evaluator)
}

fn text_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<TextValue<'a>, ScalarError>> {
    super::to_text(value, evaluator).map(Ok)
}

fn valid_serial(value: f64) -> Result<f64, ScalarError> {
    if value.is_finite() && (MIN_SERIAL..MAX_DATETIME_SERIAL).contains(&value) {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

pub(super) fn default_workday_mask() -> [bool; 7] {
    default_workdays()
}

pub(super) fn parse_date_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
) -> EvaluationResult<Result<f64, ScalarError>> {
    parser::parse(evaluator, text, parser::TextKind::Date)
        .map(|result| result.and_then(valid_serial))
}

fn parse_time_text<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
) -> EvaluationResult<Result<f64, ScalarError>> {
    parser::parse(evaluator, text, parser::TextKind::Time)
}

fn date_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match value {
        WorkingValue::Number(value) => Ok(valid_serial(value)),
        WorkingValue::Logical(value) => Ok(Ok(if value { 1.0 } else { 0.0 })),
        WorkingValue::Text(text) => parse_date_text(evaluator, text.text.as_ref()),
        WorkingValue::Error(error) => Ok(Err(error)),
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
    }
}

fn time_argument<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match value {
        WorkingValue::Number(value) => Ok(finite_time(value)),
        WorkingValue::Logical(value) => Ok(Ok(if value { 1.0 } else { 0.0 })),
        WorkingValue::Text(text) => parse_time_text(evaluator, text.text.as_ref()),
        WorkingValue::Error(error) => Ok(Err(error)),
        WorkingValue::Complex(_) => Ok(Err(ScalarError::Value)),
    }
}

fn date_only<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<CivilDate, ScalarError>> {
    match date_argument(evaluator, value)? {
        Ok(value) => Ok(serial_to_date(value)),
        Err(error) => Ok(Err(error)),
    }
}

fn finite_time(value: f64) -> Result<f64, ScalarError> {
    value
        .is_finite()
        .then_some(value)
        .ok_or(ScalarError::Number)
}

fn time_fraction(value: f64) -> f64 {
    let fraction = value - value.floor();
    if fraction >= 1.0 {
        // A negative subnormal can round to 1 when its floor is subtracted.
        // Preserve the half-open day interval used by component extraction.
        f64::from_bits(1.0_f64.to_bits() - 1)
    } else if fraction == 0.0 {
        0.0
    } else {
        fraction
    }
}

fn finish_three_number<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let first = required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value));
    let second = required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value));
    let third = required(args, 2).unwrap_or(WorkingValue::Error(ScalarError::Value));
    let first = integer_argument(evaluator, first)?;
    let second = integer_argument(evaluator, second)?;
    let third = integer_argument(evaluator, third)?;
    let result = match (first, second, third) {
        (Err(error), _, _) | (_, Err(error), _) | (_, _, Err(error)) => Err(error),
        (Ok(year), Ok(month), Ok(day)) => match i32::try_from(year) {
            Ok(year) => kernel::normalize_ymd(year, month, day).and_then(date_to_serial),
            Err(_) => Err(ScalarError::Number),
        },
    };
    push_result(evaluator, result)
}

fn apply_date<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    finish_three_number(evaluator, args)
}

fn apply_datedif<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let start = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let end = date_only(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let format = match text_argument(
        evaluator,
        required(args, 2).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )? {
        Ok(text) => text,
        Err(error) => return push_result(evaluator, Err(error)),
    };
    let format = format
        .text
        .trim_matches(|character: char| character.is_ascii_whitespace());
    let format = if format.eq_ignore_ascii_case("Y") {
        "Y"
    } else if format.eq_ignore_ascii_case("M") {
        "M"
    } else if format.eq_ignore_ascii_case("D") {
        "D"
    } else if format.eq_ignore_ascii_case("MD") {
        "MD"
    } else if format.eq_ignore_ascii_case("YM") {
        "YM"
    } else if format.eq_ignore_ascii_case("YD") {
        "YD"
    } else {
        return push_result(evaluator, Err(ScalarError::Value));
    };
    let result = match (start, end) {
        (Ok(start), Ok(end)) => kernel::datedif(start, end, format),
        (Err(error), _) | (_, Err(error)) => Err(error),
    };
    push_result(evaluator, result.map(|value| value as f64))
}

fn apply_datevalue<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let value = match text_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )? {
        Ok(text) => parse_date_text(evaluator, text.text.as_ref())?,
        Err(error) => Err(error),
    };
    push_result(
        evaluator,
        value.and_then(|serial| serial_to_date(serial).and_then(date_to_serial)),
    )
}

fn apply_day_month_year<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
    which: Function,
) -> EvaluationResult<()> {
    let date = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let result = match date {
        Ok(date) => Ok(match which {
            Function::Day => f64::from(date.day),
            Function::Month => f64::from(date.month),
            Function::Year => f64::from(date.year),
            _ => {
                return Err(EvaluationFailure::InvalidExpression(
                    "invalid date component selector",
                ));
            },
        }),
        Err(error) => Err(error),
    };
    push_result(evaluator, result)
}

fn apply_days<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let end = date_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let start = date_argument(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let result = match (end, start) {
        (Ok(end), Ok(start)) => {
            let result = end - start;
            result
                .is_finite()
                .then_some(result)
                .ok_or(ScalarError::Number)
        },
        (Err(error), _) | (_, Err(error)) => Err(error),
    };
    push_result(evaluator, result)
}

fn apply_days360<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let start = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let end = date_only(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let method = match optional(args, 2) {
        None => Ok(false),
        Some(value) => logical_argument(evaluator, value)?,
    };
    let result = match (start, end, method) {
        (Err(error), _, _) | (_, Err(error), _) | (_, _, Err(error)) => Err(error),
        (Ok(start), Ok(end), Ok(method)) => {
            if method {
                kernel::days360_european(start, end)
            } else {
                kernel::days360_us(start, end)
            }
        },
    };
    push_result(evaluator, result.map(|value| value as f64))
}

fn apply_easter<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let year = match optional(args, 0) {
        None => {
            let stamp = evaluator.context.options().calculation_timestamp().ok_or(
                EvaluationFailure::Unsupported(UnsupportedKind::CalculationClock),
            )?;
            let date = match serial_to_date(stamp.serial()) {
                Ok(date) => date,
                Err(error) => return push_result(evaluator, Err(error)),
            };
            let result = easter_sunday(date.year).and_then(|easter| {
                if kernel::date_lt(easter, date) {
                    easter_sunday(date.year + 1).and_then(date_to_serial)
                } else {
                    date_to_serial(easter)
                }
            });
            return push_result(evaluator, result);
        },
        Some(value) => match integer_argument(evaluator, value)? {
            Ok(value) => match i32::try_from(value) {
                Ok(value) => value,
                Err(_) => return push_result(evaluator, Err(ScalarError::Number)),
            },
            Err(error) => return push_result(evaluator, Err(error)),
        },
    };
    push_result(evaluator, easter_sunday(year).and_then(date_to_serial))
}

fn apply_edate<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
    end_of_month: bool,
) -> EvaluationResult<()> {
    let start = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let months = integer_argument(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let result = match (start, months) {
        (Err(error), _) | (_, Err(error)) => Err(error),
        (Ok(start), Ok(months)) => match add_months(start, months) {
            Ok(target) => {
                let target = if end_of_month {
                    calendar_date(
                        target.year,
                        target.month,
                        days_in_month(target.year, target.month),
                    )
                } else {
                    Ok(target)
                };
                target.and_then(date_to_serial)
            },
            Err(error) => Err(error),
        },
    };
    push_result(evaluator, result)
}

fn calendar_date(year: i32, month: u32, day: u32) -> Result<CivilDate, ScalarError> {
    CivilDate::new(year, month, day).ok_or(ScalarError::Number)
}

fn apply_hour_minute_second<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
    which: Function,
) -> EvaluationResult<()> {
    let time = time_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let result = match time {
        Err(error) => Err(error),
        Ok(time) => {
            if which == Function::Hour {
                Ok((time_fraction(time) * 24.0).floor())
            } else {
                let total = time * 86_400.0;
                let rounded = match super::rounding::round_nearest(total, 0.0) {
                    Ok(value) => value,
                    Err(error) => return push_result(evaluator, Err(error)),
                };
                let seconds = rounded.rem_euclid(86_400.0);
                Ok(match which {
                    Function::Minute => (seconds / 60.0).floor() % 60.0,
                    Function::Second => seconds % 60.0,
                    _ => {
                        return Err(EvaluationFailure::InvalidExpression(
                            "invalid time component selector",
                        ));
                    },
                })
            }
        },
    };
    push_result(evaluator, result)
}

fn apply_time<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let hours = number_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let minutes = number_argument(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let seconds = number_argument(
        evaluator,
        required(args, 2).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let result = match (hours, minutes, seconds) {
        (Err(error), _, _) | (_, Err(error), _) | (_, _, Err(error)) => Err(error),
        (Ok(hours), Ok(minutes), Ok(seconds)) => {
            let value = (hours * 3_600.0 + minutes * 60.0 + seconds) / 86_400.0;
            value
                .is_finite()
                .then_some(value)
                .ok_or(ScalarError::Number)
        },
    };
    push_result(evaluator, result)
}

fn apply_timevalue<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let value = match text_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )? {
        Ok(text) => parse_time_text(evaluator, text.text.as_ref())?,
        Err(error) => Err(error),
    };
    push_result(evaluator, value.and_then(finite_time))
}

fn apply_weekday<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let date = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let kind = match optional(args, 1) {
        None => Ok(1_i64),
        Some(value) => integer_argument(evaluator, value)?,
    };
    let result = match (date, kind) {
        (Err(error), _) | (_, Err(error)) => Err(error),
        (Ok(date), Ok(kind)) => weekday_result(date, kind),
    };
    push_result(evaluator, result.map(|value| value as f64))
}

fn weekday_result(date: CivilDate, kind: i64) -> Result<i64, ScalarError> {
    let sunday = i64::from(weekday_sunday_zero(date)?);
    match kind {
        1 => Ok(sunday + 1),
        2 | 11 => Ok((sunday + 6) % 7 + 1),
        3 => Ok((sunday + 6) % 7),
        12..=16 => Ok((sunday + 7 - (kind - 10)) % 7 + 1),
        17 => Ok(sunday + 1),
        _ => Err(ScalarError::Number),
    }
}

fn apply_weeknum<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
    iso_only: bool,
) -> EvaluationResult<()> {
    let date = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    if iso_only {
        return push_result(
            evaluator,
            date.and_then(|date| iso_week(date).map(|(_, week)| week as f64)),
        );
    }
    let mode = match optional(args, 1) {
        None => Ok(1_i64),
        Some(value) => number_argument(evaluator, value)?.and_then(|mode| {
            if matches!(
                mode,
                1.0 | 2.0 | 11.0 | 12.0 | 13.0 | 14.0 | 15.0 | 16.0 | 17.0 | 21.0 | 150.0
            ) {
                Ok(mode as i64)
            } else {
                Err(ScalarError::Number)
            }
        }),
    };
    let result = match (date, mode) {
        (Err(error), _) | (_, Err(error)) => Err(error),
        (Ok(date), Ok(mode)) => weeknum_result(date, mode),
    };
    push_result(evaluator, result.map(|value| value as f64))
}

fn weeknum_result(date: CivilDate, mode: i64) -> Result<i64, ScalarError> {
    if mode == 21 || mode == 150 {
        return Ok(i64::from(iso_week(date)?.1));
    }
    let start_weekday = match mode {
        1 | 17 => 0,
        2 | 11 => 1,
        12 => 2,
        13 => 3,
        14 => 4,
        15 => 5,
        16 => 6,
        _ => return Err(ScalarError::Number),
    };
    let jan1 = calendar_date(date.year, 1, 1)?;
    let day_of_year = date_difference(jan1, date)?;
    let jan1_weekday = weekday_sunday_zero(jan1)?;
    let offset = (7 + i64::from(jan1_weekday) - i64::from(start_weekday)) % 7;
    Ok((day_of_year + offset) / 7 + 1)
}

fn default_workdays() -> [bool; 7] {
    [true, false, false, false, false, false, true]
}

fn scalar_holiday<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<i64, ScalarError>> {
    match number_argument(evaluator, value)? {
        Ok(value) => Ok(serial_to_date(value)
            .and_then(date_to_serial)
            .map(|value| value as i64)),
        Err(error) => Ok(Err(error)),
    }
}

fn scalar_workdays<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    value: WorkingValue<'a>,
) -> EvaluationResult<Result<[bool; 7], ScalarError>> {
    let _ = logical_argument(evaluator, value)?;
    Ok(Err(ScalarError::Value))
}

fn networkdays_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let first = date_only(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let second = date_only(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let holiday = match optional(args, 2) {
        None => Ok(None),
        Some(value) => scalar_holiday(evaluator, value).map(|result| result.map(Some))?,
    };
    let workdays = match optional(args, 3) {
        None => Ok(default_workdays()),
        Some(value) => scalar_workdays(evaluator, value)?,
    };
    if let Some(error) = args.formula_error {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let result = match (first, second, holiday, workdays) {
        (Err(error), _, _, _)
        | (_, Err(error), _, _)
        | (_, _, Err(error), _)
        | (_, _, _, Err(error)) => Err(error),
        (Ok(first), Ok(second), Ok(holiday), Ok(workdays)) => {
            let mut holidays = [0_i64; 1];
            let holiday_count = holiday
                .map(|value| {
                    holidays[0] = value;
                    1
                })
                .unwrap_or(0);
            networkdays_loop(
                evaluator,
                first,
                second,
                &workdays,
                &holidays[..holiday_count],
            )?
        },
    };
    push_result(evaluator, result.map(|value| value as f64))
}

fn networkdays_loop<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    first: CivilDate,
    second: CivilDate,
    workdays: &[bool; 7],
    holidays: &[i64],
) -> EvaluationResult<Result<i64, ScalarError>> {
    let (start, end, sign) = if kernel::date_le(first, second) {
        (first, second, 1_i64)
    } else {
        (second, first, -1_i64)
    };
    let mut current = start;
    let mut result = 0_i64;
    let day_work = 1 + u64::from(usize::BITS - holidays.len().leading_zeros());
    loop {
        evaluator.charge_work(day_work)?;
        let is_workday = match is_workday(current, workdays, holidays) {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        if is_workday {
            result = match result.checked_add(1) {
                Some(value) => value,
                None => return Ok(Err(ScalarError::Number)),
            };
        }
        if current == end {
            break;
        }
        let next_days = match days_from_civil(current)
            .and_then(|value| value.checked_add(1).ok_or(ScalarError::Number))
        {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        current = match super::calendar::civil_from_days_checked(next_days) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };
    }
    Ok(result.checked_mul(sign).ok_or(ScalarError::Number))
}

fn workday_scalar<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    args: &mut Arguments<'a>,
) -> EvaluationResult<()> {
    let start_value = date_argument(
        evaluator,
        required(args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let offset = number_argument(
        evaluator,
        required(args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
    )?;
    let holiday = match optional(args, 2) {
        None => Ok(None),
        Some(value) => scalar_holiday(evaluator, value).map(|result| result.map(Some))?,
    };
    let workdays = match optional(args, 3) {
        None => Ok(default_workdays()),
        Some(value) => scalar_workdays(evaluator, value)?,
    };
    if let Some(error) = args.formula_error {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    let result = match (start_value, offset, holiday, workdays) {
        (Err(error), _, _, _)
        | (_, Err(error), _, _)
        | (_, _, Err(error), _)
        | (_, _, _, Err(error)) => Err(error),
        (Ok(start_value), Ok(offset), Ok(holiday), Ok(workdays)) => {
            let mut holidays = [0_i64; 1];
            let holiday_count = holiday
                .map(|value| {
                    holidays[0] = value;
                    1
                })
                .unwrap_or(0);
            workday_loop(
                evaluator,
                start_value,
                offset,
                &workdays,
                &holidays[..holiday_count],
            )?
        },
    };
    push_result(evaluator, result)
}

fn workday_loop<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    start_value: f64,
    offset: f64,
    workdays: &[bool; 7],
    holidays: &[i64],
) -> EvaluationResult<Result<f64, ScalarError>> {
    if !offset.is_finite() {
        return Ok(Err(ScalarError::Number));
    }
    let whole = offset.trunc();
    if whole == 0.0 {
        return Ok(valid_serial(start_value));
    }
    if workdays.iter().all(|off| *off) {
        return Ok(Err(ScalarError::Number));
    }
    const I64_LIMIT: f64 = 9_223_372_036_854_775_808.0;
    if whole <= -I64_LIMIT || whole >= I64_LIMIT {
        return Ok(Err(ScalarError::Number));
    }
    let whole = whole as i64;
    let steps = whole.unsigned_abs();
    let direction = if whole.is_negative() { -1_i64 } else { 1_i64 };
    let parts = match serial_to_parts(start_value) {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    let mut current = parts.date;
    let mut remaining = steps;
    let day_work = 1 + u64::from(usize::BITS - holidays.len().leading_zeros());
    while remaining != 0 {
        evaluator.charge_work(day_work)?;
        let day = match days_from_civil(current)
            .and_then(|value| value.checked_add(direction).ok_or(ScalarError::Number))
        {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        current = match super::calendar::civil_from_days_checked(day) {
            Some(value) => value,
            None => return Ok(Err(ScalarError::Number)),
        };
        let is_workday = match is_workday(current, workdays, holidays) {
            Ok(value) => value,
            Err(error) => return Ok(Err(error)),
        };
        if is_workday {
            remaining -= 1;
        }
    }
    let serial = match date_to_serial(current) {
        Ok(value) => value + parts.fraction,
        Err(error) => return Ok(Err(error)),
    };
    Ok(valid_serial(serial))
}

/// Finish an admitted sequence after its caller has scanned and normalized
/// holidays and the workweek. Both evaluation paths use these charged loops.
pub(super) fn finish_sequence(
    evaluator: &mut Evaluator<'_, '_, '_>,
    function: Function,
    first: f64,
    second: f64,
    workdays: &[bool; 7],
    holidays: &[i64],
) -> EvaluationResult<Result<f64, ScalarError>> {
    match function {
        Function::Networkdays => {
            let first = match serial_to_date(first) {
                Ok(date) => date,
                Err(error) => return Ok(Err(error)),
            };
            let second = match serial_to_date(second) {
                Ok(date) => date,
                Err(error) => return Ok(Err(error)),
            };
            networkdays_loop(evaluator, first, second, workdays, holidays)
                .map(|result| result.map(|count| count as f64))
        },
        Function::Workday => workday_loop(evaluator, first, second, workdays, holidays),
        _ => Ok(Err(ScalarError::Value)),
    }
}

fn apply_now_today<'a>(evaluator: &mut Evaluator<'a, '_, '_>, today: bool) -> EvaluationResult<()> {
    let stamp = evaluator.context.options().calculation_timestamp().ok_or(
        EvaluationFailure::Unsupported(UnsupportedKind::CalculationClock),
    )?;
    let value = if today {
        stamp.date_serial()
    } else {
        stamp.serial()
    };
    push_result(evaluator, Ok(value))
}

pub(super) fn apply<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    node: Node<'a>,
    name: &str,
) -> EvaluationResult<()> {
    let function = Function::from_name(name)
        .ok_or(EvaluationFailure::Unsupported(UnsupportedKind::Function))?;
    if !function.valid_arity(node.child_count()) {
        return evaluator.finish_invalid_arity(node);
    }
    let mut args = collect_arguments(evaluator, node)?;
    if !matches!(function, Function::Networkdays | Function::Workday)
        && let Some(error) = args.formula_error
    {
        return evaluator.push_value(WorkingValue::Error(error));
    }
    match function {
        Function::Date => apply_date(evaluator, &mut args),
        Function::DateDif => apply_datedif(evaluator, &mut args),
        Function::DateValue => apply_datevalue(evaluator, &mut args),
        Function::Day | Function::Month | Function::Year => {
            apply_day_month_year(evaluator, &mut args, function)
        },
        Function::Days => apply_days(evaluator, &mut args),
        Function::Days360 => apply_days360(evaluator, &mut args),
        Function::EasterSunday => apply_easter(evaluator, &mut args),
        Function::EDate => apply_edate(evaluator, &mut args, false),
        Function::EOMonth => apply_edate(evaluator, &mut args, true),
        Function::Hour | Function::Minute | Function::Second => {
            apply_hour_minute_second(evaluator, &mut args, function)
        },
        Function::IsoWeeknum => apply_weeknum(evaluator, &mut args, true),
        Function::Networkdays => networkdays_scalar(evaluator, &mut args),
        Function::Now => apply_now_today(evaluator, false),
        Function::Time => apply_time(evaluator, &mut args),
        Function::TimeValue => apply_timevalue(evaluator, &mut args),
        Function::Today => apply_now_today(evaluator, true),
        Function::Weekday => apply_weekday(evaluator, &mut args),
        Function::Weeknum => apply_weeknum(evaluator, &mut args, false),
        Function::Workday => workday_scalar(evaluator, &mut args),
        Function::Yearfrac => {
            let start = date_only(
                evaluator,
                required(&mut args, 0).unwrap_or(WorkingValue::Error(ScalarError::Value)),
            )?;
            let end = date_only(
                evaluator,
                required(&mut args, 1).unwrap_or(WorkingValue::Error(ScalarError::Value)),
            )?;
            let basis = match optional(&mut args, 2) {
                None => Ok(0_i64),
                Some(value) => integer_argument(evaluator, value)?,
            };
            let result = match (start, end, basis) {
                (Err(error), _, _) | (_, Err(error), _) | (_, _, Err(error)) => Err(error),
                (Ok(start), Ok(end), Ok(basis)) => yearfrac(start, end, basis),
            };
            push_result(evaluator, result)
        },
    }
}
