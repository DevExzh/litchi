//! Conversion between public table-cell formats and native format archives.

use crate::protobuf::tsk::FormatStructArchive;
use crate::{Error, Result};
use litchi_numbers::cell::data_format::control::DisplayFormat;
use litchi_numbers::cell::data_format::duration::{Style, Unit, UnitRange, Units};
use litchi_numbers::cell::data_format::number::{
    CurrencyCode, CurrencyStyle, DecimalPlaces, FixedDecimalPlaces, FractionAccuracy,
    NegativeStyle, ThousandsSeparator,
};
use litchi_numbers::cell::data_format::numeral_system::{
    Base, NegativeStyle as NumeralSystemNegativeStyle, Places,
};
use litchi_numbers::cell::data_format::{
    Checkbox, Currency, DataFormat, DateTime, Duration, Fraction, Number, NumeralSystem,
    Percentage, Scientific, StarRating, Text,
};

pub(super) const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;
pub(super) const NATIVE_CURRENCY_FORMAT_TYPE: u32 = 257;
pub(super) const NATIVE_PERCENTAGE_FORMAT_TYPE: u32 = 258;
pub(super) const NATIVE_SCIENTIFIC_FORMAT_TYPE: u32 = 259;
pub(super) const NATIVE_TEXT_FORMAT_TYPE: u32 = 260;
pub(super) const NATIVE_DATE_TIME_FORMAT_TYPE: u32 = 261;
pub(super) const NATIVE_FRACTION_FORMAT_TYPE: u32 = 262;
pub(super) const NATIVE_CHECKBOX_FORMAT_TYPE: u32 = 263;
pub(super) const NATIVE_STAR_RATING_FORMAT_TYPE: u32 = 267;
pub(super) const NATIVE_DURATION_FORMAT_TYPE: u32 = 268;
pub(super) const NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE: u32 = 269;
pub(super) const NATIVE_AUTOMATIC_DECIMAL_PLACES: u32 = 253;
const NATIVE_FRACTION_UP_TO_ONE_DIGIT: i32 = -1;
const NATIVE_FRACTION_UP_TO_TWO_DIGITS: i32 = -2;
const NATIVE_FRACTION_UP_TO_THREE_DIGITS: i32 = -3;
const NATIVE_FRACTION_HALVES: i32 = 2;
const NATIVE_FRACTION_QUARTERS: i32 = 4;
const NATIVE_FRACTION_EIGHTHS: i32 = 8;
const NATIVE_FRACTION_SIXTEENTHS: i32 = 16;
const NATIVE_FRACTION_TENTHS: i32 = 10;
const NATIVE_FRACTION_HUNDREDTHS: i32 = 100;

pub(super) fn data_format_to_native(format: &DataFormat) -> Result<FormatStructArchive> {
    if matches!(format, DataFormat::Text(_) | DataFormat::PopUpMenu(_)) {
        return Ok(text_format_to_native());
    }
    if let DataFormat::Slider(format) = format {
        return numeric_control_display_to_native(format.display_format());
    }
    if let DataFormat::Stepper(format) = format {
        return numeric_control_display_to_native(format.display_format());
    }
    if matches!(format, DataFormat::Checkbox(_)) {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_CHECKBOX_FORMAT_TYPE),
            ..Default::default()
        });
    }
    if matches!(format, DataFormat::StarRating(_)) {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_STAR_RATING_FORMAT_TYPE),
            ..Default::default()
        });
    }
    if let DataFormat::Duration(format) = format {
        let range = format.units().range();
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_DURATION_FORMAT_TYPE),
            duration_style: Some(duration_style_to_native(format.style())),
            duration_unit_largest: Some(duration_unit_to_native(range.largest())),
            duration_unit_smallest: Some(duration_unit_to_native(range.smallest())),
            use_automatic_duration_units: Some(format.units().is_automatic()),
            ..Default::default()
        });
    }
    if let DataFormat::DateTime(format) = format {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_DATE_TIME_FORMAT_TYPE),
            date_time_format: Some(format.pattern().to_owned()),
            ..Default::default()
        });
    }
    if let DataFormat::NumeralSystem(format) = format {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE),
            base: Some(u32::from(format.base().value())),
            base_places: Some(match format.places() {
                Places::Minimum => 0,
                Places::Fixed(places) => u32::from(places.value()),
            }),
            base_use_minus_sign: Some(matches!(
                format.negative_style(),
                NumeralSystemNegativeStyle::MinusSign
            )),
            ..Default::default()
        });
    }
    if let DataFormat::Fraction(format) = format {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_FRACTION_FORMAT_TYPE),
            fraction_accuracy: Some(fraction_accuracy_to_native(format.accuracy())),
            ..Default::default()
        });
    }
    let (
        format_type,
        decimal_places,
        negative_style,
        thousands_separator,
        currency_code,
        accounting_style,
    ) = match format {
        DataFormat::Automatic => {
            return Err(Error::InvalidFormat(
                "Automatic table-cell data format has no explicit payload".to_owned(),
            ));
        },
        DataFormat::Number(format) => (
            NATIVE_NUMBER_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            None,
            None,
        ),
        DataFormat::Currency(format) => (
            NATIVE_CURRENCY_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            Some(format.code().as_str().to_owned()),
            Some(matches!(format.style(), CurrencyStyle::Accounting)),
        ),
        DataFormat::Percentage(format) => (
            NATIVE_PERCENTAGE_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            None,
            None,
        ),
        DataFormat::Scientific(format) => (
            NATIVE_SCIENTIFIC_FORMAT_TYPE,
            DecimalPlaces::Fixed(format.decimal_places()),
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
            None,
            None,
        ),
        DataFormat::Fraction(_) => unreachable!("handled above"),
        DataFormat::NumeralSystem(_) => unreachable!("handled above"),
        DataFormat::DateTime(_) => unreachable!("handled above"),
        DataFormat::Duration(_) => unreachable!("handled above"),
        DataFormat::Checkbox(_) => unreachable!("handled above"),
        DataFormat::StarRating(_) => unreachable!("handled above"),
        DataFormat::Slider(_) => unreachable!("handled above"),
        DataFormat::Stepper(_) => unreachable!("handled above"),
        DataFormat::PopUpMenu(_) => unreachable!("handled above"),
        DataFormat::Text(_) => unreachable!("handled above"),
        DataFormat::Custom(_) => unreachable!("handled by custom-format registry"),
    };
    Ok(FormatStructArchive {
        format_type: Some(format_type),
        decimal_places: Some(match decimal_places {
            DecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
            DecimalPlaces::Fixed(places) => u32::from(places.value()),
        }),
        negative_style: Some(match negative_style {
            NegativeStyle::MinusSign => 0,
            NegativeStyle::Red => 1,
            NegativeStyle::Parentheses => 2,
            NegativeStyle::RedParentheses => 3,
        }),
        show_thousands_separator: Some(matches!(thousands_separator, ThousandsSeparator::Shown)),
        currency_code,
        use_accounting_style: accounting_style,
        ..Default::default()
    })
}

pub(super) fn text_format_to_native() -> FormatStructArchive {
    FormatStructArchive {
        format_type: Some(NATIVE_TEXT_FORMAT_TYPE),
        ..Default::default()
    }
}

pub(super) fn validate_text_format(native: &FormatStructArchive) -> Result<()> {
    if native != &text_format_to_native() {
        return Err(Error::InvalidFormat(
            "Pop-Up Menu references a non-canonical Text format".to_owned(),
        ));
    }
    Ok(())
}

pub(super) fn numeric_control_display_to_native(
    display: &DisplayFormat,
) -> Result<FormatStructArchive> {
    let format = match display {
        DisplayFormat::Number(format) => DataFormat::Number(*format),
        DisplayFormat::Currency(format) => DataFormat::Currency(*format),
        DisplayFormat::Percentage(format) => DataFormat::Percentage(*format),
        DisplayFormat::Fraction(format) => DataFormat::Fraction(*format),
        DisplayFormat::Scientific(format) => DataFormat::Scientific(*format),
        DisplayFormat::NumeralSystem(format) => DataFormat::NumeralSystem(*format),
    };
    data_format_to_native(&format)
}

pub(super) fn numeric_control_display_from_native(
    native: &FormatStructArchive,
) -> Result<DisplayFormat> {
    match data_format_from_native(native)? {
        DataFormat::Number(format) => Ok(DisplayFormat::Number(format)),
        DataFormat::Currency(format) => Ok(DisplayFormat::Currency(format)),
        DataFormat::Percentage(format) => Ok(DisplayFormat::Percentage(format)),
        DataFormat::Fraction(format) => Ok(DisplayFormat::Fraction(format)),
        DataFormat::Scientific(format) => Ok(DisplayFormat::Scientific(format)),
        DataFormat::NumeralSystem(format) => Ok(DisplayFormat::NumeralSystem(format)),
        _ => Err(Error::InvalidFormat(
            "Interactive numeric control uses a non-numeric display format".to_owned(),
        )),
    }
}

pub(super) fn data_format_from_native(native: &FormatStructArchive) -> Result<DataFormat> {
    let format_type = native.format_type.ok_or_else(|| {
        Error::InvalidFormat("Table cell has no native data-format type".to_owned())
    })?;
    if format_type == NATIVE_TEXT_FORMAT_TYPE {
        validate_text_format(native)?;
        return Ok(DataFormat::Text(Text));
    }
    if format_type == NATIVE_CHECKBOX_FORMAT_TYPE {
        let canonical = data_format_to_native(&DataFormat::Checkbox(Checkbox))?;
        if native != &canonical {
            return Err(Error::InvalidFormat(
                "Checkbox cell format contains non-canonical options".to_owned(),
            ));
        }
        return Ok(DataFormat::Checkbox(Checkbox));
    }
    if format_type == NATIVE_STAR_RATING_FORMAT_TYPE {
        let canonical = data_format_to_native(&DataFormat::StarRating(StarRating))?;
        if native != &canonical {
            return Err(Error::InvalidFormat(
                "Star Rating cell format contains non-canonical options".to_owned(),
            ));
        }
        return Ok(DataFormat::StarRating(StarRating));
    }
    if format_type == NATIVE_DURATION_FORMAT_TYPE {
        let style = duration_style_from_native(native.duration_style.ok_or_else(|| {
            Error::InvalidFormat("Duration cell format has no presentation style".to_owned())
        })?)?;
        let range = UnitRange::new(
            duration_unit_from_native(native.duration_unit_largest.ok_or_else(|| {
                Error::InvalidFormat("Duration cell format has no largest unit".to_owned())
            })?)?,
            duration_unit_from_native(native.duration_unit_smallest.ok_or_else(|| {
                Error::InvalidFormat("Duration cell format has no smallest unit".to_owned())
            })?)?,
        )?;
        let units = if native.use_automatic_duration_units.ok_or_else(|| {
            Error::InvalidFormat("Duration cell format has no unit-selection mode".to_owned())
        })? {
            Units::Automatic(range)
        } else {
            Units::Custom(range)
        };
        let format = Duration::new(style, units);
        let canonical = data_format_to_native(&DataFormat::Duration(format))?;
        if native != &canonical {
            return Err(Error::InvalidFormat(
                "Duration cell format contains non-canonical options".to_owned(),
            ));
        }
        return Ok(DataFormat::Duration(format));
    }
    if format_type == NATIVE_DATE_TIME_FORMAT_TYPE {
        if native.decimal_places.is_some()
            || native.negative_style.is_some()
            || native.show_thousands_separator.is_some()
            || native.currency_code.is_some()
            || native.use_accounting_style.is_some()
            || native.fraction_accuracy.is_some()
            || native.base.is_some()
            || native.base_places.is_some()
            || native.base_use_minus_sign.is_some()
            || native.suppress_date_format.is_some()
            || native.suppress_time_format.is_some()
        {
            return Err(Error::InvalidFormat(
                "Date & Time cell format contains non-canonical options".to_owned(),
            ));
        }
        let pattern = native.date_time_format.as_deref().ok_or_else(|| {
            Error::InvalidFormat("Date & Time cell format has no ICU pattern".to_owned())
        })?;
        return DateTime::new(pattern)
            .map(DataFormat::DateTime)
            .map_err(Into::into);
    }
    if format_type == NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE {
        if native.decimal_places.is_some()
            || native.negative_style.is_some()
            || native.show_thousands_separator.is_some()
            || native.currency_code.is_some()
            || native.use_accounting_style.is_some()
            || native.fraction_accuracy.is_some()
            || native.date_time_format.is_some()
            || native.suppress_date_format.is_some()
            || native.suppress_time_format.is_some()
        {
            return Err(Error::InvalidFormat(
                "Numeral System cell format contains non-canonical numeric options".to_owned(),
            ));
        }
        let base = native.base.ok_or_else(|| {
            Error::InvalidFormat("Numeral System cell format has no base".to_owned())
        })?;
        let base = u8::try_from(base).map_err(|_| {
            Error::InvalidFormat(format!(
                "Numeral System cell format has invalid base {base}"
            ))
        })?;
        let base = Base::new(base)?;
        let places = native.base_places.ok_or_else(|| {
            Error::InvalidFormat("Numeral System cell format has no places setting".to_owned())
        })?;
        let places = match places {
            0 => Places::Minimum,
            value => {
                let value = u8::try_from(value).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "Numeral System cell format has invalid places setting {value}"
                    ))
                })?;
                Places::fixed(value)?
            },
        };
        let negative_style = match native.base_use_minus_sign {
            Some(true) => NumeralSystemNegativeStyle::MinusSign,
            Some(false) => NumeralSystemNegativeStyle::TwosComplement,
            None => {
                return Err(Error::InvalidFormat(
                    "Numeral System cell format has no negative-style setting".to_owned(),
                ));
            },
        };
        return NumeralSystem::new(base, places, negative_style)
            .map(DataFormat::NumeralSystem)
            .map_err(Into::into);
    }
    if format_type == NATIVE_FRACTION_FORMAT_TYPE {
        let accuracy = native.fraction_accuracy.ok_or_else(|| {
            Error::InvalidFormat("Fraction cell format has no accuracy setting".to_owned())
        })?;
        if native.decimal_places.is_some()
            || native.negative_style.is_some()
            || native.show_thousands_separator.is_some()
            || native.currency_code.is_some()
            || native.use_accounting_style.is_some()
            || native.base.is_some()
            || native.base_places.is_some()
            || native.base_use_minus_sign.is_some()
            || native.date_time_format.is_some()
            || native.suppress_date_format.is_some()
            || native.suppress_time_format.is_some()
        {
            return Err(Error::InvalidFormat(
                "Fraction cell format contains non-canonical decimal options".to_owned(),
            ));
        }
        return Ok(DataFormat::Fraction(Fraction::new(
            fraction_accuracy_from_native(accuracy)?,
        )));
    }
    if !matches!(
        format_type,
        NATIVE_NUMBER_FORMAT_TYPE
            | NATIVE_CURRENCY_FORMAT_TYPE
            | NATIVE_PERCENTAGE_FORMAT_TYPE
            | NATIVE_SCIENTIFIC_FORMAT_TYPE
    ) {
        return Err(Error::InvalidFormat(format!(
            "Table cell uses unsupported native data-format type {format_type}"
        )));
    }
    if native.fraction_accuracy.is_some() {
        return Err(Error::InvalidFormat(
            "Non-Fraction cell format contains a fraction-accuracy setting".to_owned(),
        ));
    }
    if native.base.is_some() || native.base_places.is_some() || native.base_use_minus_sign.is_some()
    {
        return Err(Error::InvalidFormat(
            "Non-Numeral-System cell format contains numeral-system options".to_owned(),
        ));
    }
    if native.date_time_format.is_some()
        || native.suppress_date_format.is_some()
        || native.suppress_time_format.is_some()
    {
        return Err(Error::InvalidFormat(
            "Non-Date-Time cell format contains Date & Time options".to_owned(),
        ));
    }
    let decimal_places = match native.decimal_places {
        Some(NATIVE_AUTOMATIC_DECIMAL_PLACES) => DecimalPlaces::Automatic,
        Some(value) => {
            let value = u8::try_from(value).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Table cell has invalid decimal-place count {value}"
                ))
            })?;
            DecimalPlaces::Fixed(FixedDecimalPlaces::new(value)?)
        },
        None => {
            return Err(Error::InvalidFormat(
                "Table-cell data format has no decimal-place setting".to_owned(),
            ));
        },
    };
    let negative_style = match native.negative_style {
        Some(0) => NegativeStyle::MinusSign,
        Some(1) => NegativeStyle::Red,
        Some(2) => NegativeStyle::Parentheses,
        Some(3) => NegativeStyle::RedParentheses,
        value => {
            return Err(Error::InvalidFormat(format!(
                "Table cell has invalid negative-number style {value:?}"
            )));
        },
    };
    let thousands_separator = match native.show_thousands_separator {
        Some(false) => ThousandsSeparator::Hidden,
        Some(true) => ThousandsSeparator::Shown,
        None => {
            return Err(Error::InvalidFormat(
                "Table-cell data format has no thousands-separator setting".to_owned(),
            ));
        },
    };
    Ok(match format_type {
        NATIVE_NUMBER_FORMAT_TYPE => DataFormat::Number(Number::new(
            decimal_places,
            negative_style,
            thousands_separator,
        )),
        NATIVE_CURRENCY_FORMAT_TYPE => {
            let code = native.currency_code.as_deref().ok_or_else(|| {
                Error::InvalidFormat("Currency cell format has no currency code".to_owned())
            })?;
            let currency_code = CurrencyCode::new(code)?;
            let style = match native.use_accounting_style {
                Some(false) => CurrencyStyle::Standard,
                Some(true) => CurrencyStyle::Accounting,
                None => {
                    return Err(Error::InvalidFormat(
                        "Currency cell format has no accounting-style setting".to_owned(),
                    ));
                },
            };
            DataFormat::Currency(Currency::new(
                currency_code,
                decimal_places,
                negative_style,
                thousands_separator,
                style,
            ))
        },
        NATIVE_PERCENTAGE_FORMAT_TYPE => DataFormat::Percentage(Percentage::new(
            decimal_places,
            negative_style,
            thousands_separator,
        )),
        NATIVE_SCIENTIFIC_FORMAT_TYPE => {
            if negative_style != NegativeStyle::MinusSign
                || thousands_separator != ThousandsSeparator::Hidden
            {
                return Err(Error::InvalidFormat(
                    "Scientific cell format contains non-canonical decimal options".to_owned(),
                ));
            }
            let DecimalPlaces::Fixed(decimal_places) = decimal_places else {
                return Err(Error::InvalidFormat(
                    "Scientific cell format cannot use automatic decimal places".to_owned(),
                ));
            };
            DataFormat::Scientific(Scientific::new(decimal_places))
        },
        _ => unreachable!("validated native data-format type"),
    })
}

const fn duration_unit_to_native(unit: Unit) -> u32 {
    match unit {
        Unit::Weeks => 1,
        Unit::Days => 2,
        Unit::Hours => 4,
        Unit::Minutes => 8,
        Unit::Seconds => 16,
        Unit::Milliseconds => 32,
    }
}

fn duration_unit_from_native(value: u32) -> Result<Unit> {
    match value {
        1 => Ok(Unit::Weeks),
        2 => Ok(Unit::Days),
        4 => Ok(Unit::Hours),
        8 => Ok(Unit::Minutes),
        16 => Ok(Unit::Seconds),
        32 => Ok(Unit::Milliseconds),
        _ => Err(Error::InvalidFormat(format!(
            "Unsupported native Duration unit {value}"
        ))),
    }
}

const fn duration_style_to_native(style: Style) -> u32 {
    match style {
        Style::Colon => 0,
        Style::Abbreviated => 1,
        Style::FullNames => 2,
    }
}

fn duration_style_from_native(value: u32) -> Result<Style> {
    match value {
        0 => Ok(Style::Colon),
        1 => Ok(Style::Abbreviated),
        2 => Ok(Style::FullNames),
        _ => Err(Error::InvalidFormat(format!(
            "Unsupported native Duration style {value}"
        ))),
    }
}

const fn fraction_accuracy_to_native(accuracy: FractionAccuracy) -> u32 {
    match accuracy {
        FractionAccuracy::UpToOneDigit => NATIVE_FRACTION_UP_TO_ONE_DIGIT as u32,
        FractionAccuracy::UpToTwoDigits => NATIVE_FRACTION_UP_TO_TWO_DIGITS as u32,
        FractionAccuracy::UpToThreeDigits => NATIVE_FRACTION_UP_TO_THREE_DIGITS as u32,
        FractionAccuracy::Halves => NATIVE_FRACTION_HALVES as u32,
        FractionAccuracy::Quarters => NATIVE_FRACTION_QUARTERS as u32,
        FractionAccuracy::Eighths => NATIVE_FRACTION_EIGHTHS as u32,
        FractionAccuracy::Sixteenths => NATIVE_FRACTION_SIXTEENTHS as u32,
        FractionAccuracy::Tenths => NATIVE_FRACTION_TENTHS as u32,
        FractionAccuracy::Hundredths => NATIVE_FRACTION_HUNDREDTHS as u32,
    }
}

fn fraction_accuracy_from_native(value: u32) -> Result<FractionAccuracy> {
    match value as i32 {
        NATIVE_FRACTION_UP_TO_ONE_DIGIT => Ok(FractionAccuracy::UpToOneDigit),
        NATIVE_FRACTION_UP_TO_TWO_DIGITS => Ok(FractionAccuracy::UpToTwoDigits),
        NATIVE_FRACTION_UP_TO_THREE_DIGITS => Ok(FractionAccuracy::UpToThreeDigits),
        NATIVE_FRACTION_HALVES => Ok(FractionAccuracy::Halves),
        NATIVE_FRACTION_QUARTERS => Ok(FractionAccuracy::Quarters),
        NATIVE_FRACTION_EIGHTHS => Ok(FractionAccuracy::Eighths),
        NATIVE_FRACTION_SIXTEENTHS => Ok(FractionAccuracy::Sixteenths),
        NATIVE_FRACTION_TENTHS => Ok(FractionAccuracy::Tenths),
        NATIVE_FRACTION_HUNDREDTHS => Ok(FractionAccuracy::Hundredths),
        _ => Err(Error::InvalidFormat(format!(
            "Fraction cell format has invalid accuracy {value}"
        ))),
    }
}
