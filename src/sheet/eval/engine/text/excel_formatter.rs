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
