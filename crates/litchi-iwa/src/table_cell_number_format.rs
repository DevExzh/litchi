//! Private conversion seam for the extracted table-cell number vocabulary.

use litchi_iwa_common::table::cell::number_format::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, NumberFormat, ThousandsSeparator,
};

use crate::table_cell_data_format::{
    TableCellDecimalPlaces, TableCellFixedDecimalPlaces, TableCellNegativeNumberStyle,
    TableCellNumberFormat, TableCellThousandsSeparator,
};
use crate::{Error, Result};

impl From<litchi_iwa_common::table::cell::number_format::Error> for Error {
    fn from(error: litchi_iwa_common::table::cell::number_format::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

pub(crate) fn from_native(format: TableCellNumberFormat) -> Result<NumberFormat> {
    let decimal_places = match format.decimal_places() {
        TableCellDecimalPlaces::Automatic => DecimalPlaces::Automatic,
        TableCellDecimalPlaces::Fixed(value) => {
            DecimalPlaces::Fixed(FixedDecimalPlaces::new(value.value())?)
        },
    };
    Ok(NumberFormat::new(
        decimal_places,
        match format.negative_style() {
            TableCellNegativeNumberStyle::MinusSign => NegativeStyle::MinusSign,
            TableCellNegativeNumberStyle::Red => NegativeStyle::Red,
            TableCellNegativeNumberStyle::Parentheses => NegativeStyle::Parentheses,
            TableCellNegativeNumberStyle::RedParentheses => NegativeStyle::RedParentheses,
        },
        match format.thousands_separator() {
            TableCellThousandsSeparator::Hidden => ThousandsSeparator::Hidden,
            TableCellThousandsSeparator::Shown => ThousandsSeparator::Shown,
        },
    ))
}

pub(crate) fn to_native(format: NumberFormat) -> Result<TableCellNumberFormat> {
    let decimal_places = match format.decimal_places() {
        DecimalPlaces::Automatic => TableCellDecimalPlaces::Automatic,
        DecimalPlaces::Fixed(value) => {
            TableCellDecimalPlaces::Fixed(TableCellFixedDecimalPlaces::new(value.value())?)
        },
    };
    Ok(TableCellNumberFormat::new(
        decimal_places,
        match format.negative_style() {
            NegativeStyle::MinusSign => TableCellNegativeNumberStyle::MinusSign,
            NegativeStyle::Red => TableCellNegativeNumberStyle::Red,
            NegativeStyle::Parentheses => TableCellNegativeNumberStyle::Parentheses,
            NegativeStyle::RedParentheses => TableCellNegativeNumberStyle::RedParentheses,
        },
        match format.thousands_separator() {
            ThousandsSeparator::Hidden => TableCellThousandsSeparator::Hidden,
            ThousandsSeparator::Shown => TableCellThousandsSeparator::Shown,
        },
    ))
}
