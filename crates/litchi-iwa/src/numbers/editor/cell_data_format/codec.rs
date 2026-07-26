//! Conversion between public table-cell formats and native format archives.

use crate::protobuf::tsk::FormatStructArchive;
use crate::table_cell_data_format::{
    TableCellCurrencyCode, TableCellCurrencyFormat, TableCellCurrencyStyle, TableCellDataFormat,
    TableCellDecimalPlaces, TableCellFixedDecimalPlaces, TableCellFractionAccuracy,
    TableCellFractionFormat, TableCellNegativeNumberStyle, TableCellNumberFormat,
    TableCellNumeralSystemBase, TableCellNumeralSystemFormat, TableCellNumeralSystemNegativeStyle,
    TableCellNumeralSystemPlaces, TableCellPercentageFormat, TableCellScientificFormat,
    TableCellThousandsSeparator,
};
use crate::{Error, Result};

pub(super) const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;
pub(super) const NATIVE_CURRENCY_FORMAT_TYPE: u32 = 257;
pub(super) const NATIVE_PERCENTAGE_FORMAT_TYPE: u32 = 258;
pub(super) const NATIVE_SCIENTIFIC_FORMAT_TYPE: u32 = 259;
pub(super) const NATIVE_FRACTION_FORMAT_TYPE: u32 = 262;
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

pub(super) fn data_format_to_native(format: TableCellDataFormat) -> Result<FormatStructArchive> {
    if let TableCellDataFormat::NumeralSystem(format) = format {
        return Ok(FormatStructArchive {
            format_type: Some(NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE),
            base: Some(u32::from(format.base().value())),
            base_places: Some(match format.places() {
                TableCellNumeralSystemPlaces::Minimum => 0,
                TableCellNumeralSystemPlaces::Fixed(places) => u32::from(places.value()),
            }),
            base_use_minus_sign: Some(matches!(
                format.negative_style(),
                TableCellNumeralSystemNegativeStyle::MinusSign
            )),
            ..Default::default()
        });
    }
    if let TableCellDataFormat::Fraction(format) = format {
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
        TableCellDataFormat::Automatic => {
            return Err(Error::InvalidFormat(
                "Automatic table-cell data format has no explicit payload".to_owned(),
            ));
        },
        TableCellDataFormat::Number(format) => (
            NATIVE_NUMBER_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            None,
            None,
        ),
        TableCellDataFormat::Currency(format) => (
            NATIVE_CURRENCY_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            Some(format.currency_code().as_str().to_owned()),
            Some(matches!(format.style(), TableCellCurrencyStyle::Accounting)),
        ),
        TableCellDataFormat::Percentage(format) => (
            NATIVE_PERCENTAGE_FORMAT_TYPE,
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            None,
            None,
        ),
        TableCellDataFormat::Scientific(format) => (
            NATIVE_SCIENTIFIC_FORMAT_TYPE,
            TableCellDecimalPlaces::Fixed(format.decimal_places()),
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
            None,
            None,
        ),
        TableCellDataFormat::Fraction(_) => unreachable!("handled above"),
        TableCellDataFormat::NumeralSystem(_) => unreachable!("handled above"),
    };
    Ok(FormatStructArchive {
        format_type: Some(format_type),
        decimal_places: Some(match decimal_places {
            TableCellDecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
            TableCellDecimalPlaces::Fixed(places) => u32::from(places.value()),
        }),
        negative_style: Some(match negative_style {
            TableCellNegativeNumberStyle::MinusSign => 0,
            TableCellNegativeNumberStyle::Red => 1,
            TableCellNegativeNumberStyle::Parentheses => 2,
            TableCellNegativeNumberStyle::RedParentheses => 3,
        }),
        show_thousands_separator: Some(matches!(
            thousands_separator,
            TableCellThousandsSeparator::Shown
        )),
        currency_code,
        use_accounting_style: accounting_style,
        ..Default::default()
    })
}

pub(super) fn data_format_from_native(native: &FormatStructArchive) -> Result<TableCellDataFormat> {
    let format_type = native.format_type.ok_or_else(|| {
        Error::InvalidFormat("Table cell has no native data-format type".to_owned())
    })?;
    if format_type == NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE {
        if native.decimal_places.is_some()
            || native.negative_style.is_some()
            || native.show_thousands_separator.is_some()
            || native.currency_code.is_some()
            || native.use_accounting_style.is_some()
            || native.fraction_accuracy.is_some()
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
        let base = TableCellNumeralSystemBase::new(base)?;
        let places = native.base_places.ok_or_else(|| {
            Error::InvalidFormat("Numeral System cell format has no places setting".to_owned())
        })?;
        let places = match places {
            0 => TableCellNumeralSystemPlaces::Minimum,
            value => {
                let value = u8::try_from(value).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "Numeral System cell format has invalid places setting {value}"
                    ))
                })?;
                TableCellNumeralSystemPlaces::fixed(value)?
            },
        };
        let negative_style = match native.base_use_minus_sign {
            Some(true) => TableCellNumeralSystemNegativeStyle::MinusSign,
            Some(false) => TableCellNumeralSystemNegativeStyle::TwosComplement,
            None => {
                return Err(Error::InvalidFormat(
                    "Numeral System cell format has no negative-style setting".to_owned(),
                ));
            },
        };
        return TableCellNumeralSystemFormat::new(base, places, negative_style)
            .map(TableCellDataFormat::NumeralSystem);
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
        {
            return Err(Error::InvalidFormat(
                "Fraction cell format contains non-canonical decimal options".to_owned(),
            ));
        }
        return Ok(TableCellDataFormat::Fraction(TableCellFractionFormat::new(
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
    let decimal_places = match native.decimal_places {
        Some(NATIVE_AUTOMATIC_DECIMAL_PLACES) => TableCellDecimalPlaces::Automatic,
        Some(value) => {
            let value = u8::try_from(value).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Table cell has invalid decimal-place count {value}"
                ))
            })?;
            TableCellDecimalPlaces::Fixed(TableCellFixedDecimalPlaces::new(value)?)
        },
        None => {
            return Err(Error::InvalidFormat(
                "Table-cell data format has no decimal-place setting".to_owned(),
            ));
        },
    };
    let negative_style = match native.negative_style {
        Some(0) => TableCellNegativeNumberStyle::MinusSign,
        Some(1) => TableCellNegativeNumberStyle::Red,
        Some(2) => TableCellNegativeNumberStyle::Parentheses,
        Some(3) => TableCellNegativeNumberStyle::RedParentheses,
        value => {
            return Err(Error::InvalidFormat(format!(
                "Table cell has invalid negative-number style {value:?}"
            )));
        },
    };
    let thousands_separator = match native.show_thousands_separator {
        Some(false) => TableCellThousandsSeparator::Hidden,
        Some(true) => TableCellThousandsSeparator::Shown,
        None => {
            return Err(Error::InvalidFormat(
                "Table-cell data format has no thousands-separator setting".to_owned(),
            ));
        },
    };
    Ok(match format_type {
        NATIVE_NUMBER_FORMAT_TYPE => TableCellDataFormat::Number(TableCellNumberFormat::new(
            decimal_places,
            negative_style,
            thousands_separator,
        )),
        NATIVE_CURRENCY_FORMAT_TYPE => {
            let code = native.currency_code.as_deref().ok_or_else(|| {
                Error::InvalidFormat("Currency cell format has no currency code".to_owned())
            })?;
            let currency_code = TableCellCurrencyCode::new(code)?;
            let style = match native.use_accounting_style {
                Some(false) => TableCellCurrencyStyle::Standard,
                Some(true) => TableCellCurrencyStyle::Accounting,
                None => {
                    return Err(Error::InvalidFormat(
                        "Currency cell format has no accounting-style setting".to_owned(),
                    ));
                },
            };
            TableCellDataFormat::Currency(TableCellCurrencyFormat::new(
                currency_code,
                decimal_places,
                negative_style,
                thousands_separator,
                style,
            ))
        },
        NATIVE_PERCENTAGE_FORMAT_TYPE => TableCellDataFormat::Percentage(
            TableCellPercentageFormat::new(decimal_places, negative_style, thousands_separator),
        ),
        NATIVE_SCIENTIFIC_FORMAT_TYPE => {
            if negative_style != TableCellNegativeNumberStyle::MinusSign
                || thousands_separator != TableCellThousandsSeparator::Hidden
            {
                return Err(Error::InvalidFormat(
                    "Scientific cell format contains non-canonical decimal options".to_owned(),
                ));
            }
            let TableCellDecimalPlaces::Fixed(decimal_places) = decimal_places else {
                return Err(Error::InvalidFormat(
                    "Scientific cell format cannot use automatic decimal places".to_owned(),
                ));
            };
            TableCellDataFormat::Scientific(TableCellScientificFormat::new(decimal_places))
        },
        _ => unreachable!("validated native data-format type"),
    })
}

const fn fraction_accuracy_to_native(accuracy: TableCellFractionAccuracy) -> u32 {
    match accuracy {
        TableCellFractionAccuracy::UpToOneDigit => NATIVE_FRACTION_UP_TO_ONE_DIGIT as u32,
        TableCellFractionAccuracy::UpToTwoDigits => NATIVE_FRACTION_UP_TO_TWO_DIGITS as u32,
        TableCellFractionAccuracy::UpToThreeDigits => NATIVE_FRACTION_UP_TO_THREE_DIGITS as u32,
        TableCellFractionAccuracy::Halves => NATIVE_FRACTION_HALVES as u32,
        TableCellFractionAccuracy::Quarters => NATIVE_FRACTION_QUARTERS as u32,
        TableCellFractionAccuracy::Eighths => NATIVE_FRACTION_EIGHTHS as u32,
        TableCellFractionAccuracy::Sixteenths => NATIVE_FRACTION_SIXTEENTHS as u32,
        TableCellFractionAccuracy::Tenths => NATIVE_FRACTION_TENTHS as u32,
        TableCellFractionAccuracy::Hundredths => NATIVE_FRACTION_HUNDREDTHS as u32,
    }
}

fn fraction_accuracy_from_native(value: u32) -> Result<TableCellFractionAccuracy> {
    match value as i32 {
        NATIVE_FRACTION_UP_TO_ONE_DIGIT => Ok(TableCellFractionAccuracy::UpToOneDigit),
        NATIVE_FRACTION_UP_TO_TWO_DIGITS => Ok(TableCellFractionAccuracy::UpToTwoDigits),
        NATIVE_FRACTION_UP_TO_THREE_DIGITS => Ok(TableCellFractionAccuracy::UpToThreeDigits),
        NATIVE_FRACTION_HALVES => Ok(TableCellFractionAccuracy::Halves),
        NATIVE_FRACTION_QUARTERS => Ok(TableCellFractionAccuracy::Quarters),
        NATIVE_FRACTION_EIGHTHS => Ok(TableCellFractionAccuracy::Eighths),
        NATIVE_FRACTION_SIXTEENTHS => Ok(TableCellFractionAccuracy::Sixteenths),
        NATIVE_FRACTION_TENTHS => Ok(TableCellFractionAccuracy::Tenths),
        NATIVE_FRACTION_HUNDREDTHS => Ok(TableCellFractionAccuracy::Hundredths),
        _ => Err(Error::InvalidFormat(format!(
            "Fraction cell format has invalid accuracy {value}"
        ))),
    }
}
