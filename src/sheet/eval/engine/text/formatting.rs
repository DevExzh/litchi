use super::excel_formatter::{
    CellFormat, FormattedData, detect_custom_number_format, format_excel_f64,
};
use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_bool, to_number, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use chrono::Datelike;

const MAX_FIXED_DECIMALS: i32 = 30;

pub(crate) async fn eval_fixed(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 3 {
        return Ok(CellValue::Error(
            "FIXED expects 1 to 3 arguments (number, [decimals], [no_commas])".to_string(),
        ));
    }

    let number_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&number_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("FIXED number is not numeric".to_string())),
    };

    let decimals = if args.len() >= 2 {
        let dec_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&dec_val) {
            Some(n) => n.trunc() as i32,
            None => {
                return Ok(CellValue::Error(
                    "FIXED decimals is not numeric".to_string(),
                ));
            },
        }
    } else {
        2
    };
    let decimals = decimals.clamp(-MAX_FIXED_DECIMALS, MAX_FIXED_DECIMALS);

    let no_commas = if args.len() == 3 {
        let flag_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        to_bool(&flag_val)
    } else {
        false
    };

    let rounded = round_to_decimal_places(number, decimals);
    let display_decimals = decimals.max(0) as usize;
    let formatted = format_number_text(rounded, display_decimals, !no_commas);
    Ok(CellValue::String(formatted))
}

pub(crate) async fn eval_dollar(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "DOLLAR expects 1 or 2 arguments (number, [decimals])".to_string(),
        ));
    }

    let number_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let number = match to_number(&number_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("DOLLAR number is not numeric".to_string())),
    };

    let decimals = if args.len() == 2 {
        let dec_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&dec_val) {
            Some(n) => n.trunc() as i32,
            None => {
                return Ok(CellValue::Error(
                    "DOLLAR decimals is not numeric".to_string(),
                ));
            },
        }
    } else {
        2
    };
    let decimals = decimals.clamp(-MAX_FIXED_DECIMALS, MAX_FIXED_DECIMALS);

    let rounded = round_to_decimal_places(number, decimals);
    let display_decimals = decimals.max(0) as usize;
    let text = format_currency_text(rounded, display_decimals, "$");
    Ok(CellValue::String(text))
}

pub(crate) async fn eval_text(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "TEXT expects 2 arguments (value, format_text)".to_string(),
        ));
    }

    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let format_expr = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let format_text = to_text(&format_expr);
    if format_text.is_empty() {
        return Ok(CellValue::Error(
            "TEXT format_text must not be empty".to_string(),
        ));
    }

    let formatted = format_with_pattern(ctx, &value, &format_text);
    Ok(formatted)
}

fn format_with_pattern(ctx: EvalCtx<'_>, value: &CellValue, pattern: &str) -> CellValue {
    let sections: Vec<&str> = pattern.split(';').collect();

    match value {
        CellValue::String(s) => {
            if sections.len() >= 4 {
                // The 4th section is for text. Replace @ with the string.
                let text_format = sections[3];
                let mut result = String::new();
                let mut chars = text_format.chars().peekable();
                while let Some(c) = chars.next() {
                    match c {
                        '\\' => {
                            if let Some(next) = chars.next() {
                                result.push(next);
                            }
                        },
                        '"' => {
                            for next in chars.by_ref() {
                                if next == '"' {
                                    break;
                                }
                                result.push(next);
                            }
                        },
                        '@' => {
                            result.push_str(s);
                        },
                        _ => result.push(c),
                    }
                }
                CellValue::String(result)
            } else {
                // If no 4th section, return as-is
                CellValue::String(s.clone())
            }
        },
        CellValue::Bool(_) => CellValue::Error("#VALUE!".to_string()),
        CellValue::Empty => {
            // Excel TREATS empty as 0 for TEXT function
            format_number_with_pattern(ctx, 0.0, pattern)
        },
        CellValue::Error(err) => CellValue::Error(err.clone()),
        CellValue::Formula { .. } => {
            // In evaluation engine, we should ideally have the cached result or evaluated value
            // If we still have a Formula here, something is wrong or it's a raw value
            CellValue::Error("#VALUE!".to_string())
        },
        _ => {
            if let Some(number) = to_number(value) {
                format_number_with_pattern(ctx, number, pattern)
            } else {
                CellValue::Error("#VALUE!".to_string())
            }
        },
    }
}

fn format_number_with_pattern(ctx: EvalCtx<'_>, number: f64, pattern: &str) -> CellValue {
    let sections: Vec<&str> = pattern.split(';').collect();
    let format_to_use = if sections.len() > 1 {
        if number > 0.0 {
            sections[0]
        } else if number < 0.0 {
            sections[1]
        } else {
            sections.get(2).unwrap_or(&sections[0])
        }
    } else {
        pattern
    };

    let cell_format = detect_custom_number_format(format_to_use);
    let is_1904 = ctx.is_1904_date_system();
    let data_ref = format_excel_f64(number, Some(&cell_format), is_1904);

    match data_ref {
        FormattedData::DateTime(dt) => {
            let (y, m, d, hh, mins, ss, _) = dt.to_ymd_hms_milli();
            let _is_duration = cell_format == CellFormat::TimeDelta;

            let mut result = String::new();
            let mut chars = format_to_use.chars().peekable();
            let mut has_h = false;

            while let Some(c) = chars.next() {
                match c {
                    '\\' => {
                        if let Some(next) = chars.next() {
                            result.push(next);
                        }
                    },
                    '"' => {
                        for next in chars.by_ref() {
                            if next == '"' {
                                break;
                            }
                            result.push(next);
                        }
                    },
                    '[' => {
                        let mut bracket_content = String::new();
                        for next in chars.by_ref() {
                            if next == ']' {
                                break;
                            }
                            bracket_content.push(next);
                        }
                        match bracket_content.to_lowercase().as_str() {
                            "h" => {
                                let total_hours = number * 24.0;
                                result.push_str(&format!("{}", total_hours.trunc() as i64));
                                has_h = true;
                            },
                            "m" | "mm" => {
                                let total_mins = number * 24.0 * 60.0;
                                if bracket_content.len() == 2 {
                                    result.push_str(&format!("{:02}", total_mins.trunc() as i64));
                                } else {
                                    result.push_str(&format!("{}", total_mins.trunc() as i64));
                                }
                            },
                            "s" | "ss" => {
                                let total_secs = number * 24.0 * 60.0 * 60.0;
                                if bracket_content.len() == 2 {
                                    result.push_str(&format!("{:02}", total_secs.trunc() as i64));
                                } else {
                                    result.push_str(&format!("{}", total_secs.trunc() as i64));
                                }
                            },
                            _ => {}, // Ignore colors etc for now
                        }
                    },
                    'y' | 'Y' => {
                        let mut count = 1;
                        while let Some(&next) = chars.peek() {
                            if next == 'y' || next == 'Y' {
                                count += 1;
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        if count >= 4 {
                            result.push_str(&format!("{:04}", y));
                        } else {
                            result.push_str(&format!("{:02}", y % 100));
                        }
                    },
                    'm' | 'M' => {
                        let mut count = 1;
                        while let Some(&next) = chars.peek() {
                            if next == 'm' || next == 'M' {
                                count += 1;
                                chars.next();
                            } else {
                                break;
                            }
                        }

                        // Check if it's minutes (follows h or precedes s)
                        let is_mins = has_h || matches!(chars.peek(), Some(&'s') | Some(&'S'));

                        if is_mins {
                            if count >= 2 {
                                result.push_str(&format!("{:02}", mins));
                            } else {
                                result.push_str(&format!("{}", mins));
                            }
                        } else {
                            match count {
                                1 => result.push_str(&format!("{}", m)),
                                2 => result.push_str(&format!("{:02}", m)),
                                3 => result.push_str(match m {
                                    1 => "Jan",
                                    2 => "Feb",
                                    3 => "Mar",
                                    4 => "Apr",
                                    5 => "May",
                                    6 => "Jun",
                                    7 => "Jul",
                                    8 => "Aug",
                                    9 => "Sep",
                                    10 => "Oct",
                                    11 => "Nov",
                                    _ => "Dec",
                                }),
                                _ => result.push_str(match m {
                                    1 => "January",
                                    2 => "February",
                                    3 => "March",
                                    4 => "April",
                                    5 => "May",
                                    6 => "June",
                                    7 => "July",
                                    8 => "August",
                                    9 => "September",
                                    10 => "October",
                                    11 => "November",
                                    _ => "December",
                                }),
                            }
                        }
                    },
                    'd' | 'D' => {
                        let mut count = 1;
                        while let Some(&next) = chars.peek() {
                            if next == 'd' || next == 'D' {
                                count += 1;
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        match count {
                            1 => result.push_str(&format!("{}", d)),
                            2 => result.push_str(&format!("{:02}", d)),
                            3 | 4 => {
                                // Day of week calculation
                                if let Some(date) =
                                    chrono::NaiveDate::from_ymd_opt(y as i32, m as u32, d as u32)
                                {
                                    let wd = date.weekday();
                                    if count == 3 {
                                        result.push_str(match wd {
                                            chrono::Weekday::Sun => "Sun",
                                            chrono::Weekday::Mon => "Mon",
                                            chrono::Weekday::Tue => "Tue",
                                            chrono::Weekday::Wed => "Wed",
                                            chrono::Weekday::Thu => "Thu",
                                            chrono::Weekday::Fri => "Fri",
                                            chrono::Weekday::Sat => "Sat",
                                        });
                                    } else {
                                        result.push_str(match wd {
                                            chrono::Weekday::Sun => "Sunday",
                                            chrono::Weekday::Mon => "Monday",
                                            chrono::Weekday::Tue => "Tuesday",
                                            chrono::Weekday::Wed => "Wednesday",
                                            chrono::Weekday::Thu => "Thursday",
                                            chrono::Weekday::Fri => "Friday",
                                            chrono::Weekday::Sat => "Saturday",
                                        });
                                    }
                                }
                            },
                            _ => {},
                        }
                    },
                    'h' | 'H' => {
                        let mut count = 1;
                        while let Some(&next) = chars.peek() {
                            if next == 'h' || next == 'H' {
                                count += 1;
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        let mut display_h = hh;
                        let is_12h = format_to_use.to_lowercase().contains("am/pm");
                        if is_12h {
                            display_h = if hh == 0 {
                                12
                            } else if hh > 12 {
                                hh - 12
                            } else {
                                hh
                            };
                        }
                        if count >= 2 {
                            result.push_str(&format!("{:02}", display_h));
                        } else {
                            result.push_str(&format!("{}", display_h));
                        }
                        has_h = true;
                    },
                    's' | 'S' => {
                        let mut count = 1;
                        while let Some(&next) = chars.peek() {
                            if next == 's' || next == 'S' {
                                count += 1;
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        if count >= 2 {
                            result.push_str(&format!("{:02}", ss));
                        } else {
                            result.push_str(&format!("{}", ss));
                        }
                    },
                    'a' | 'A' => {
                        // Check for AM/PM
                        let mut matched = false;
                        if let Some(&'m') | Some(&'M') = chars.peek() {
                            // Potentially AM/PM
                            let remaining: String = chars.clone().take(4).collect();
                            if remaining.to_lowercase().starts_with("m/pm") {
                                for _ in 0..4 {
                                    chars.next();
                                }
                                let is_pm = hh >= 12;
                                if c.is_uppercase() {
                                    result.push_str(if is_pm { "PM" } else { "AM" });
                                } else {
                                    result.push_str(if is_pm { "pm" } else { "am" });
                                }
                                matched = true;
                            }
                        }
                        if !matched {
                            result.push(c);
                        }
                    },
                    _ => result.push(c),
                }
            }
            CellValue::String(result)
        },
        FormattedData::Float(f) => {
            // Very simple numeric formatting fallback
            if format_to_use == "0" {
                CellValue::String(format!("{:.0}", f))
            } else if format_to_use == "0.00" {
                CellValue::String(format!("{:.2}", f))
            } else {
                CellValue::String(format!("{}", f))
            }
        },
        _ => CellValue::Error("#VALUE!".to_string()),
    }
}

fn round_to_decimal_places(value: f64, decimals: i32) -> f64 {
    if decimals >= 0 {
        let factor = 10f64.powi(decimals);
        if factor.is_infinite() {
            return value;
        }
        (value * factor).round() / factor
    } else {
        let factor = 10f64.powi(-decimals);
        if factor.is_infinite() {
            return 0.0;
        }
        (value / factor).round() * factor
    }
}

fn format_number_text(value: f64, decimals: usize, use_commas: bool) -> String {
    let sign = if value.is_sign_negative() { "-" } else { "" };
    let core = format_abs_value(value.abs(), decimals, use_commas);
    format!("{sign}{core}")
}

fn format_currency_text(value: f64, decimals: usize, currency_symbol: &str) -> String {
    let sign = if value.is_sign_negative() { "-" } else { "" };
    let core = format_abs_value(value.abs(), decimals, true);
    format!("{sign}{currency_symbol}{core}")
}

fn format_abs_value(abs_value: f64, decimals: usize, use_commas: bool) -> String {
    let formatted = format!("{:.*}", decimals, abs_value);
    if !use_commas {
        return formatted;
    }

    let mut parts = formatted.splitn(2, '.');
    let int_part = parts.next().unwrap_or("");
    let frac_part = parts.next();
    let int_with_commas = insert_commas(int_part);

    if let Some(frac) = frac_part {
        if decimals > 0 {
            format!("{}.{}", int_with_commas, frac)
        } else {
            int_with_commas
        }
    } else {
        int_with_commas
    }
}

fn insert_commas(digits: &str) -> String {
    if digits.len() <= 3 {
        return digits.to_string();
    }
    let mut result = String::with_capacity(digits.len() + digits.len() / 3);
    let chars: Vec<char> = digits.chars().collect();
    for (idx, ch) in chars.iter().enumerate() {
        if idx > 0 && (chars.len() - idx).is_multiple_of(3) {
            result.push(',');
        }
        result.push(*ch);
    }
    result
}
