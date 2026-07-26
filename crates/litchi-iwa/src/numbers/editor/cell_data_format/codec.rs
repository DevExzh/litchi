//! Conversion between public table-cell formats and native format archives.

use crate::protobuf::tsk::FormatStructArchive;
use crate::table_cell_data_format::{
    TableCellCurrencyCode, TableCellCurrencyFormat, TableCellCurrencyStyle, TableCellDataFormat,
    TableCellDecimalPlaces, TableCellFixedDecimalPlaces, TableCellNegativeNumberStyle,
    TableCellNumberFormat, TableCellPercentageFormat, TableCellScientificFormat,
    TableCellThousandsSeparator,
};
use crate::{Error, Result};

pub(super) const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;
pub(super) const NATIVE_CURRENCY_FORMAT_TYPE: u32 = 257;
pub(super) const NATIVE_PERCENTAGE_FORMAT_TYPE: u32 = 258;
pub(super) const NATIVE_SCIENTIFIC_FORMAT_TYPE: u32 = 259;
pub(super) const NATIVE_AUTOMATIC_DECIMAL_PLACES: u32 = 253;

pub(super) fn data_format_to_native(format: TableCellDataFormat) -> Result<FormatStructArchive> {
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
