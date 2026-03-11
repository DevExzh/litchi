// Ported from calamine (MIT License)

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ExcelDateTimeType {
    DateTime,
    TimeDelta,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ExcelDateTime {
    pub value: f64,
    pub datetime_type: ExcelDateTimeType,
    pub is_1904: bool,
}

impl ExcelDateTime {
    pub fn new(value: f64, datetime_type: ExcelDateTimeType, is_1904: bool) -> Self {
        ExcelDateTime {
            value,
            datetime_type,
            is_1904,
        }
    }

    pub fn to_ymd_hms_milli(self) -> (u16, u8, u8, u8, u8, u8, u16) {
        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        let mut days = self.value.floor() as u64;

        if self.is_1904 {
            days += 111_033;
        } else if days > 365 {
            days += 109_571;
        } else {
            days += 109_572;
        }

        let year_days_400 = days / 146097;
        let mut days = days % 146097;

        let year_days_100;
        if days < 36525 {
            year_days_100 = days / 36525;
            days %= 36525;
        } else {
            year_days_100 = 1 + (days - 36525) / 36524;
            days = (days - 36525) % 36524;
        }

        let year_days_4;
        let mut non_leap_year_block = false;
        if year_days_100 == 0 {
            year_days_4 = days / 1461;
            days %= 1461;
        } else if days < 1460 {
            year_days_4 = days / 1460;
            days %= 1460;
            non_leap_year_block = true;
        } else {
            year_days_4 = 1 + (days - 1460) / 1461;
            days = (days - 1460) % 1461;
        }

        let year_days_1;
        if non_leap_year_block {
            year_days_1 = days / 365;
            days %= 365;
        } else if days < 366 {
            year_days_1 = days / 366;
            days %= 366;
        } else {
            year_days_1 = 1 + (days - 366) / 365;
            days = (days - 366) % 365;
        }

        let year = 1600 + year_days_400 * 400 + year_days_100 * 100 + year_days_4 * 4 + year_days_1;
        days += 1;

        if year.is_multiple_of(4) && (!year.is_multiple_of(100) || year.is_multiple_of(400)) {
            months[1] = 29;
        }

        if !self.is_1904 && year == 1900 {
            months[1] = 29;
            if self.value.trunc() == 366.0 {
                days += 1;
            }
        }

        let mut month = 1;
        for month_days in months {
            if days > month_days {
                days -= month_days;
                month += 1;
            } else {
                break;
            }
        }

        let day = days;
        let time = self.value.fract();
        let day_seconds = 24.0 * 60.0 * 60.0;
        let milli = ((time * day_seconds).fract() * 1000.0).round() as u64;
        let day_as_seconds = (time * day_seconds) as u64;

        let hour = day_as_seconds / 3600;
        let min = (day_as_seconds - hour * 3600) / 60;
        let sec = (day_as_seconds - hour * 3600 - min * 60) % 60;

        (
            year as u16,
            month as u8,
            day as u8,
            hour as u8,
            min as u8,
            sec as u8,
            milli as u16,
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(dead_code)]
pub enum FormattedData {
    Int(i64),
    Float(f64),
    DateTime(ExcelDateTime),
}

pub fn detect_custom_number_format(format: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;
    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false,
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return CellFormat::Other,
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta,
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return CellFormat::DateTime,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return CellFormat::DateTime;
            },
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            },
        }
        prev = s;
    }
    CellFormat::Other
}

pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> FormattedData {
    match format {
        Some(CellFormat::DateTime) => FormattedData::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => FormattedData::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => FormattedData::Float(value),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ===== CellFormat enum tests =====

    #[test]
    fn test_cell_format_equality() {
        assert_eq!(CellFormat::Other, CellFormat::Other);
        assert_eq!(CellFormat::DateTime, CellFormat::DateTime);
        assert_eq!(CellFormat::TimeDelta, CellFormat::TimeDelta);
        assert_ne!(CellFormat::Other, CellFormat::DateTime);
    }

    #[test]
    fn test_excel_datetime_type_equality() {
        assert_eq!(ExcelDateTimeType::DateTime, ExcelDateTimeType::DateTime);
        assert_eq!(ExcelDateTimeType::TimeDelta, ExcelDateTimeType::TimeDelta);
        assert_ne!(ExcelDateTimeType::DateTime, ExcelDateTimeType::TimeDelta);
    }

    // ===== ExcelDateTime::new tests =====

    #[test]
    fn test_excel_datetime_new() {
        let dt = ExcelDateTime::new(44561.5, ExcelDateTimeType::DateTime, false);
        assert_eq!(dt.value, 44561.5);
        assert_eq!(dt.datetime_type, ExcelDateTimeType::DateTime);
        assert!(!dt.is_1904);

        let dt = ExcelDateTime::new(100.0, ExcelDateTimeType::TimeDelta, true);
        assert_eq!(dt.value, 100.0);
        assert_eq!(dt.datetime_type, ExcelDateTimeType::TimeDelta);
        assert!(dt.is_1904);
    }

    // ===== ExcelDateTime::to_ymd_hms_milli tests - 1900 date system =====

    #[test]
    fn test_to_ymd_hms_milli_1900_epoch() {
        // Day 0 in 1900 system actually returns 1899-12-31 (implementation detail)
        let dt = ExcelDateTime::new(0.0, ExcelDateTimeType::DateTime, false);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 1899);
        assert_eq!(month, 12);
        assert_eq!(day, 31);
        assert_eq!(hour, 0);
        assert_eq!(min, 0);
        assert_eq!(sec, 0);
        assert_eq!(milli, 0);
    }

    #[test]
    fn test_to_ymd_hms_milli_1900_day1() {
        // Day 1 is 1900-01-01
        let dt = ExcelDateTime::new(1.0, ExcelDateTimeType::DateTime, false);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 1900);
        assert_eq!(month, 1);
        assert_eq!(day, 1);
        assert_eq!(hour, 0);
        assert_eq!(min, 0);
        assert_eq!(sec, 0);
        assert_eq!(milli, 0);
    }

    #[test]
    fn test_to_ymd_hms_milli_1900_with_time() {
        // 44561.5 = 2021-12-31 12:00:00 (noon)
        let dt = ExcelDateTime::new(44561.5, ExcelDateTimeType::DateTime, false);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 2021);
        assert_eq!(month, 12);
        assert_eq!(day, 31);
        assert_eq!(hour, 12);
        assert_eq!(min, 0);
        assert_eq!(sec, 0);
        assert_eq!(milli, 0);
    }

    #[test]
    fn test_to_ymd_hms_milli_1900_feb_1900() {
        // Excel treats 1900 as a leap year (incorrectly)
        // Day 60 is 1900-02-29 in Excel
        let dt = ExcelDateTime::new(60.0, ExcelDateTimeType::DateTime, false);
        let (year, month, day, _, _, _, _) = dt.to_ymd_hms_milli();
        assert_eq!(year, 1900);
        assert_eq!(month, 2);
        assert_eq!(day, 29);
    }

    // ===== ExcelDateTime::to_ymd_hms_milli tests - 1904 date system =====

    #[test]
    fn test_to_ymd_hms_milli_1904_epoch() {
        // 1904 system: day 0 is 1904-01-01
        let dt = ExcelDateTime::new(0.0, ExcelDateTimeType::DateTime, true);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 1904);
        assert_eq!(month, 1);
        assert_eq!(day, 1);
        assert_eq!(hour, 0);
        assert_eq!(min, 0);
        assert_eq!(sec, 0);
        assert_eq!(milli, 0);
    }

    #[test]
    fn test_to_ymd_hms_milli_1904_with_time() {
        // 44561.5 adjusted for 1904 system
        let dt = ExcelDateTime::new(44561.5, ExcelDateTimeType::DateTime, true);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 2026); // 4 years later than 1900 system
        assert_eq!(month, 1);
        assert_eq!(day, 1);
        assert_eq!(hour, 12);
        assert_eq!(min, 0);
        assert_eq!(sec, 0);
        assert_eq!(milli, 0);
    }

    #[test]
    fn test_to_ymd_hms_milli_with_milliseconds() {
        // 0.5 days = 12 hours, with fractional milliseconds
        let dt = ExcelDateTime::new(1.000011574, ExcelDateTimeType::DateTime, false);
        let (year, month, day, hour, min, sec, milli) = dt.to_ymd_hms_milli();
        assert_eq!(year, 1900);
        assert_eq!(month, 1);
        assert_eq!(day, 1);
        // Small time fraction should give us some milliseconds
        assert!(milli > 0 || hour > 0 || min > 0 || sec > 0);
    }

    // ===== detect_custom_number_format tests =====

    #[test]
    fn test_detect_custom_number_format_date() {
        assert_eq!(
            detect_custom_number_format("yyyy-mm-dd"),
            CellFormat::DateTime
        );
        assert_eq!(
            detect_custom_number_format("dd/mm/yyyy"),
            CellFormat::DateTime
        );
        assert_eq!(detect_custom_number_format("m/d/yy"), CellFormat::DateTime);
    }

    #[test]
    fn test_detect_custom_number_format_time() {
        assert_eq!(
            detect_custom_number_format("hh:mm:ss"),
            CellFormat::DateTime
        );
        assert_eq!(
            detect_custom_number_format("h:mm AM/PM"),
            CellFormat::DateTime
        );
    }

    #[test]
    fn test_detect_custom_number_format_datetime() {
        assert_eq!(
            detect_custom_number_format("yyyy-mm-dd hh:mm:ss"),
            CellFormat::DateTime
        );
    }

    #[test]
    fn test_detect_custom_number_format_timedelta() {
        assert_eq!(
            detect_custom_number_format("[h]:mm:ss"),
            CellFormat::TimeDelta
        );
        assert_eq!(
            detect_custom_number_format("[hh]:mm:ss"),
            CellFormat::TimeDelta
        );
        assert_eq!(detect_custom_number_format("[m]:ss"), CellFormat::TimeDelta);
    }

    #[test]
    fn test_detect_custom_number_format_other() {
        assert_eq!(detect_custom_number_format("0.00"), CellFormat::Other);
        assert_eq!(detect_custom_number_format("#,##0"), CellFormat::Other);
        assert_eq!(
            detect_custom_number_format("0.00%;[Red]-0.00%"),
            CellFormat::Other
        );
    }

    #[test]
    fn test_detect_custom_number_format_with_quotes() {
        // Quotes should not affect detection
        assert_eq!(
            detect_custom_number_format("\"Date: \"yyyy-mm-dd"),
            CellFormat::DateTime
        );
        assert_eq!(
            detect_custom_number_format("\"Number: \"0.00"),
            CellFormat::Other
        );
    }

    #[test]
    fn test_detect_custom_number_format_with_escapes() {
        // Escaped characters should be ignored
        assert_eq!(
            detect_custom_number_format("\\dyyyy-mm-dd"),
            CellFormat::DateTime
        );
    }

    #[test]
    fn test_detect_custom_number_format_empty() {
        assert_eq!(detect_custom_number_format(""), CellFormat::Other);
    }

    // ===== format_excel_f64 tests =====

    #[test]
    fn test_format_excel_f64_datetime() {
        let value = 44561.5;
        let format = CellFormat::DateTime;
        let result = format_excel_f64(value, Some(&format), false);
        match result {
            FormattedData::DateTime(dt) => {
                assert_eq!(dt.value, 44561.5);
                assert_eq!(dt.datetime_type, ExcelDateTimeType::DateTime);
                assert!(!dt.is_1904);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[test]
    fn test_format_excel_f64_timedelta() {
        let value = 1.5; // 1.5 days
        let format = CellFormat::TimeDelta;
        let result = format_excel_f64(value, Some(&format), false);
        match result {
            FormattedData::DateTime(dt) => {
                assert_eq!(dt.value, 1.5);
                assert_eq!(dt.datetime_type, ExcelDateTimeType::TimeDelta);
                assert!(!dt.is_1904);
            },
            _ => panic!("Expected DateTime"),
        }
    }

    #[test]
    fn test_format_excel_f64_float() {
        let value = 123.456;
        let result = format_excel_f64(value, Some(&CellFormat::Other), false);
        match result {
            FormattedData::Float(v) => assert_eq!(v, 123.456),
            _ => panic!("Expected Float"),
        }
    }

    #[test]
    fn test_format_excel_f64_none_format() {
        let value = 123.456;
        let result = format_excel_f64(value, None, false);
        match result {
            FormattedData::Float(v) => assert_eq!(v, 123.456),
            _ => panic!("Expected Float when format is None"),
        }
    }

    #[test]
    fn test_format_excel_f64_int_value() {
        // Integer values should be preserved
        let value = 100.0;
        let result = format_excel_f64(value, None, false);
        match result {
            FormattedData::Float(v) => assert_eq!(v, 100.0),
            _ => panic!("Expected Float"),
        }
    }

    // ===== FormattedData enum tests =====

    #[test]
    fn test_formatted_data_variants() {
        let int_data = FormattedData::Int(42);
        let float_data = FormattedData::Float(3.14);
        let datetime_data = FormattedData::DateTime(ExcelDateTime::new(
            44561.0,
            ExcelDateTimeType::DateTime,
            false,
        ));

        match int_data {
            FormattedData::Int(v) => assert_eq!(v, 42),
            _ => panic!("Expected Int"),
        }

        match float_data {
            FormattedData::Float(v) => assert_eq!(v, 3.14),
            _ => panic!("Expected Float"),
        }

        match datetime_data {
            FormattedData::DateTime(dt) => {
                assert_eq!(dt.value, 44561.0);
            },
            _ => panic!("Expected DateTime"),
        }
    }
}
