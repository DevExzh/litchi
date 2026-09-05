//! Private native owner for focused display-format transactions.
//!
//! Number, Percentage, Scientific, and Fraction cells use the same BNC decimal
//! cell kind and the same format-list graph. Duration, Currency, Date/Time,
//! and Text use their nominal native families while sharing the same
//! format-list graph. This module keeps that graph surgery in one route while
//! retaining a typed family boundary at every codec and semantic conversion.
//! It intentionally exposes no native identifiers or arbitrary format-type
//! input to the package API.

use std::borrow::Cow;

use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireFieldView, WireView},
};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_currency_format_codec as currency_codec,
    numbers_table_cell_date_time_format_codec as date_time_codec,
    numbers_table_cell_duration_format_codec as duration_codec,
    numbers_table_cell_fraction_format_codec as fraction_codec,
    numbers_table_cell_number_format_codec as number_codec,
    numbers_table_cell_percentage_format_codec as percentage_codec,
    numbers_table_cell_scientific_format_codec as scientific_codec,
    numbers_table_cell_storage_codec as storage_codec,
    numbers_table_cell_text_format_codec as text_codec,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind, NumericCellType};

use super::table_cell_control_native as native;
use super::{
    Package, table_cell_pop_up_menu as popup,
    table_cell_pop_up_menu::{CellTarget, Error, Path, TransactionBudget},
    table_cell_pop_up_menu_native as popup_native,
};
use crate::cell::data_format::currency::{Currency, CurrencyCode, CurrencyStyle};
use crate::cell::data_format::custom::{
    Condition, ConditionValue, Custom, DateTime as CustomDateTime, DateTimePattern, MAX_NAME_BYTES,
    MAX_PATTERN_BYTES, Name, Number as CustomNumber, NumberPattern, NumberRule, Text as CustomText,
};
use crate::cell::data_format::duration::{
    Duration, Style as DurationStyle, Unit as DurationUnit, UnitRange, Units as DurationUnits,
};
use crate::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, Fraction, FractionAccuracy, NegativeStyle, Number,
    Percentage, Scientific, ThousandsSeparator,
};
use crate::cell::data_format::{DateTime, Text};

const NUMBER_FORMAT_TYPE: u32 = number_codec::NATIVE_NUMBER_FORMAT_TYPE;
const CURRENCY_FORMAT_TYPE: u32 = currency_codec::NATIVE_CURRENCY_FORMAT_TYPE;
const PERCENTAGE_FORMAT_TYPE: u32 = percentage_codec::NATIVE_PERCENTAGE_FORMAT_TYPE;
const SCIENTIFIC_FORMAT_TYPE: u32 = scientific_codec::NATIVE_SCIENTIFIC_FORMAT_TYPE;
const FRACTION_FORMAT_TYPE: u32 = fraction_codec::NATIVE_FRACTION_FORMAT_TYPE;
const DATE_TIME_FORMAT_TYPE: u32 = date_time_codec::NATIVE_DATE_TIME_FORMAT_TYPE;
const DURATION_FORMAT_TYPE: u32 = duration_codec::NATIVE_DURATION_FORMAT_TYPE;
// Numbers stores the plain Text display format in the same
// `FormatStructArchive` family as the scalar formats.  The strict control
// codec accepts this native discriminator while this private owner keeps it
// out of the public API.
const TEXT_FORMAT_TYPE: u32 = text_codec::NATIVE_TEXT_FORMAT_TYPE;

// BNC v5's fixed-layout fields are private to `litchi-numbers-wire`.  The
// converted-text marker retains the generic Number-format key in field
// `CELL_FORMAT_IDENTIFIER_FLAG`; this local, private layout is used only to
// preserve and validate that native secondary reference.  No identifier
// crosses the Numbers package boundary.
const BNC_HEADER_LEN: usize = 12;
const BNC_EXPLICIT_FORMAT_FLAGS_START: usize = 6;
const BNC_EXPLICIT_FORMAT_FLAGS_END: usize = 8;
const BNC_CELL_FORMAT_KIND_FLAG: u32 = 0x0000_1000;
const BNC_CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_4000;
const BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_8000;
const BNC_DURATION_FORMAT_IDENTIFIER_FLAG: u32 = 0x0001_0000;
const BNC_TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG: u32 = 0x0004_0000;
const BNC_RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;
const BNC_CONTROL_CELL_SPEC_FLAG: u32 = 0x0000_0400;
const BNC_TEXT_ALLOWED_FORMAT_FLAGS: u32 =
    BNC_CELL_FORMAT_KIND_FLAG | BNC_CELL_FORMAT_IDENTIFIER_FLAG | BNC_TEXT_FORMAT_IDENTIFIER_FLAG;
const BNC_FORMAT_FLAGS: u32 = BNC_CONTROL_CELL_SPEC_FLAG
    | BNC_CELL_FORMAT_KIND_FLAG
    | BNC_CELL_FORMAT_IDENTIFIER_FLAG
    | BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG
    | BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG
    | BNC_DURATION_FORMAT_IDENTIFIER_FLAG
    | BNC_TEXT_FORMAT_IDENTIFIER_FLAG
    | BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG;
const BNC_FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0001, 16),
    (0x0000_0002, 8),
    (0x0000_0004, 8),
    (0x0000_0008, 4),
    (0x0000_0010, 4),
    (0x0000_0020, 4),
    (0x0000_0040, 4),
    (0x0000_0080, 4),
    (0x0000_0100, 4),
    (0x0000_0200, 4),
    (BNC_CONTROL_CELL_SPEC_FLAG, 4),
    (0x0000_0800, 4),
    (BNC_CELL_FORMAT_KIND_FLAG, 4),
    (BNC_CELL_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_CURRENCY_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_DATE_TIME_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_DURATION_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_TEXT_FORMAT_IDENTIFIER_FLAG, 4),
    (BNC_CHECKBOX_FORMAT_IDENTIFIER_FLAG, 4),
    (0x0008_0000, 4),
    (0x0010_0000, 4),
];

/// The only display families admitted by the focused display-format owner.
///
/// The enum is private to the package adapter.  In particular, callers cannot
/// supply an arbitrary native discriminator and use this route as a generic
/// raw-ID editor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DisplayFormatFamily {
    Number,
    Currency,
    Percentage,
    Scientific,
    Fraction,
    DateTime,
    Duration,
    Text,
}

impl DisplayFormatFamily {
    const fn native_type(self) -> u32 {
        match self {
            Self::Number => NUMBER_FORMAT_TYPE,
            Self::Currency => CURRENCY_FORMAT_TYPE,
            Self::Percentage => PERCENTAGE_FORMAT_TYPE,
            Self::Scientific => SCIENTIFIC_FORMAT_TYPE,
            Self::Fraction => FRACTION_FORMAT_TYPE,
            Self::DateTime => DATE_TIME_FORMAT_TYPE,
            Self::Duration => DURATION_FORMAT_TYPE,
            Self::Text => TEXT_FORMAT_TYPE,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DisplayReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for DisplayReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Number package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NumberFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for NumberFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Percentage package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PercentageFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for PercentageFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Currency package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CurrencyFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for CurrencyFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Scientific package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ScientificFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for ScientificFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Fraction package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FractionFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for FractionFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Date & Time package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DateTimeFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for DateTimeFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Duration package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DurationFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for DurationFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

/// Typed native failure returned to the Text package facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum TextFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for TextFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum NativeDisplayValue {
    Number(Number),
    Currency(Currency),
    Percentage(Percentage),
    Scientific(Scientific),
    Fraction(Fraction),
    DateTime(DateTime),
    Duration(Duration),
    Text(Text),
}

impl NativeDisplayValue {
    const fn family(&self) -> DisplayFormatFamily {
        match self {
            Self::Number(_) => DisplayFormatFamily::Number,
            Self::Currency(_) => DisplayFormatFamily::Currency,
            Self::Percentage(_) => DisplayFormatFamily::Percentage,
            Self::Scientific(_) => DisplayFormatFamily::Scientific,
            Self::Fraction(_) => DisplayFormatFamily::Fraction,
            Self::DateTime(_) => DisplayFormatFamily::DateTime,
            Self::Duration(_) => DisplayFormatFamily::Duration,
            Self::Text(_) => DisplayFormatFamily::Text,
        }
    }

    fn from_parts(
        family: DisplayFormatFamily,
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
        currency_code: Option<&str>,
        use_accounting_style: Option<bool>,
        path: Path,
    ) -> Result<Self, Error> {
        let decimal_places = if decimal_places == 253 {
            DecimalPlaces::Automatic
        } else {
            DecimalPlaces::Fixed(
                FixedDecimalPlaces::new(
                    u8::try_from(decimal_places).map_err(|_| Error::InvalidSource { path })?,
                )
                .map_err(|_| Error::InvalidSource { path })?,
            )
        };
        let negative_style = match negative_style {
            0 => NegativeStyle::MinusSign,
            1 => NegativeStyle::Red,
            2 => NegativeStyle::Parentheses,
            3 => NegativeStyle::RedParentheses,
            _ => return Err(Error::InvalidSource { path }),
        };
        let thousands_separator = if show_thousands_separator {
            ThousandsSeparator::Shown
        } else {
            ThousandsSeparator::Hidden
        };
        let value = match family {
            DisplayFormatFamily::Number => NativeDisplayValue::Number(Number::new(
                decimal_places,
                negative_style,
                thousands_separator,
            )),
            DisplayFormatFamily::Currency => {
                let code = currency_code
                    .ok_or(Error::InvalidSource { path })
                    .and_then(|code| {
                        CurrencyCode::new(code).map_err(|_| Error::InvalidSource { path })
                    })?;
                let style = match use_accounting_style.ok_or(Error::InvalidSource { path })? {
                    true => CurrencyStyle::Accounting,
                    false => CurrencyStyle::Standard,
                };
                NativeDisplayValue::Currency(Currency::new(
                    code,
                    decimal_places,
                    negative_style,
                    thousands_separator,
                    style,
                ))
            },
            DisplayFormatFamily::Percentage => NativeDisplayValue::Percentage(Percentage::new(
                decimal_places,
                negative_style,
                thousands_separator,
            )),
            DisplayFormatFamily::Scientific => {
                let DecimalPlaces::Fixed(decimal_places) = decimal_places else {
                    return Err(Error::InvalidSource { path });
                };
                if native_negative_style(negative_style)
                    != scientific_codec::NATIVE_SCIENTIFIC_NEGATIVE_STYLE
                    || (matches!(thousands_separator, ThousandsSeparator::Shown))
                        != scientific_codec::NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR
                {
                    return Err(Error::InvalidSource { path });
                }
                NativeDisplayValue::Scientific(Scientific::new(decimal_places))
            },
            DisplayFormatFamily::Fraction => return Err(Error::InvalidSource { path }),
            DisplayFormatFamily::DateTime => return Err(Error::InvalidSource { path }),
            DisplayFormatFamily::Duration => return Err(Error::InvalidSource { path }),
            DisplayFormatFamily::Text => return Err(Error::InvalidSource { path }),
        };
        Ok(value)
    }
}

fn rewrite_display_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_limit = source
        .len()
        .checked_add(8)
        .ok_or(Error::InvalidSource { path })?;
    let owned_cell_bytes = source
        .len()
        .checked_add(output_limit)
        .ok_or(Error::InvalidSource { path })?;
    let work = output_limit
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(source.len()))
        .ok_or(Error::InvalidSource { path })?;
    let allocations = litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS
        .checked_mul(2)
        // One encoded output and the two shared decimal-family metadata
        // buffers.
        .and_then(|amount| amount.checked_add(3))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(allocations, path)?;
    budget.charge_scratch_bytes(owned_cell_bytes, path)?;
    budget.charge_retained_bytes(output_limit, path)?;
    budget.charge_transaction_work(work, path)?;

    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_numeric_type = cell.numeric_cell_type();
    cell.set_number_or_percentage_format_identifier_preserving_value(desired_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let output = cell
        .try_encode_with_limit(output_limit)
        .map_err(|error| match error {
            litchi_numbers_wire::Error::Allocation { requested } => Error::Allocation {
                amount: requested,
                path,
            },
            litchi_numbers_wire::Error::InvalidFormat(_)
            | litchi_numbers_wire::Error::ParseError(_)
            | litchi_numbers_wire::Error::OutputLimitExceeded { .. } => {
                Error::InvalidSource { path }
            },
        })?;
    let candidate_cell = BncCell::parse(&output).map_err(|_| Error::Verification)?;
    verify_display_cell_metadata(
        source_value,
        source_cache,
        source_numeric_type,
        &candidate_cell,
        desired_identifier,
    )?;
    Ok(output)
}

fn charge_owned_bnc_parse(
    source_len: usize,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.charge_allocations(litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS, path)?;
    budget.charge_scratch_bytes(source_len, path)?;
    budget.charge_transaction_work(source_len, path)
}

fn verify_display_cell_metadata(
    source_value: litchi_numbers_wire::StoredValue,
    source_cache: Option<litchi_numbers_wire::CachedScalar>,
    source_numeric_type: Option<NumericCellType>,
    candidate_cell: &BncCell,
    desired_identifier: Option<u32>,
) -> Result<(), Error> {
    let expected_explicit =
        desired_identifier.map_or(0, |_| litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT);
    if source_value != candidate_cell.stored_value()
        || candidate_cell.cached_scalar().ok() != Some(source_cache)
    {
        return Err(Error::Verification);
    }
    if candidate_cell.explicit_format_flags() != expected_explicit
        || candidate_cell.numeric_cell_type() != source_numeric_type
        || candidate_cell.cell_format_kind()
            != desired_identifier.map(|_| litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND)
        || candidate_cell.format_identifier() != desired_identifier
        || candidate_cell.secondary_format_identifier().is_some()
        || candidate_cell.control_cell_spec_identifier().is_some()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn rewrite_date_time_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_limit = source
        .len()
        .checked_add(8)
        .ok_or(Error::InvalidSource { path })?;
    let owned_cell_bytes = source
        .len()
        .checked_add(output_limit)
        .ok_or(Error::InvalidSource { path })?;
    let work = output_limit
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(source.len()))
        .ok_or(Error::InvalidSource { path })?;
    let allocations = litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(3))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(allocations, path)?;
    budget.charge_scratch_bytes(owned_cell_bytes, path)?;
    budget.charge_retained_bytes(output_limit, path)?;
    budget.charge_transaction_work(work, path)?;

    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_numeric_type = cell.numeric_cell_type();
    if !cell.is_date_time_format_compatible() {
        return Err(Error::UnsupportedDependency { path });
    }
    cell.set_date_time_format_identifier_preserving_value(desired_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let output = cell
        .try_encode_with_limit(output_limit)
        .map_err(|error| match error {
            litchi_numbers_wire::Error::Allocation { requested } => Error::Allocation {
                amount: requested,
                path,
            },
            litchi_numbers_wire::Error::InvalidFormat(_)
            | litchi_numbers_wire::Error::ParseError(_)
            | litchi_numbers_wire::Error::OutputLimitExceeded { .. } => {
                Error::InvalidSource { path }
            },
        })?;
    let candidate_cell = BncCell::parse(&output).map_err(|_| Error::Verification)?;
    verify_date_time_cell_metadata(
        source_value,
        source_cache,
        source_numeric_type,
        &candidate_cell,
        desired_identifier,
    )?;
    Ok(output)
}

fn verify_date_time_cell_metadata(
    source_value: litchi_numbers_wire::StoredValue,
    source_cache: Option<litchi_numbers_wire::CachedScalar>,
    source_numeric_type: Option<NumericCellType>,
    candidate_cell: &BncCell,
    desired_identifier: Option<u32>,
) -> Result<(), Error> {
    let expected_explicit =
        desired_identifier.map_or(0, |_| litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT);
    if source_value != candidate_cell.stored_value()
        || candidate_cell.cached_scalar().ok() != Some(source_cache)
        || candidate_cell.numeric_cell_type() != source_numeric_type
        || !candidate_cell.is_date_time_format_compatible()
        || !candidate_cell.has_only_date_time_format_metadata()
    {
        return Err(Error::Verification);
    }
    if candidate_cell.explicit_format_flags() != expected_explicit
        || candidate_cell.cell_format_kind()
            != desired_identifier.map(|_| litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND)
        || candidate_cell.format_identifier() != desired_identifier
        || candidate_cell.secondary_format_identifier().is_some()
        || candidate_cell.control_cell_spec_identifier().is_some()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn rewrite_duration_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_limit = source
        .len()
        .checked_add(12)
        .ok_or(Error::InvalidSource { path })?;
    let owned_cell_bytes = source
        .len()
        .checked_add(output_limit)
        .ok_or(Error::InvalidSource { path })?;
    let work = output_limit
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(source.len()))
        .ok_or(Error::InvalidSource { path })?;
    let allocations = litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS
        .checked_mul(2)
        // One encoded output plus the three small format-field buffers that
        // the Duration mutation can materialize.
        .and_then(|amount| amount.checked_add(4))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(allocations, path)?;
    budget.charge_scratch_bytes(owned_cell_bytes, path)?;
    budget.charge_retained_bytes(output_limit, path)?;
    budget.charge_transaction_work(work, path)?;

    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_secondary = cell.secondary_format_identifier();
    if !cell.is_duration_format_compatible() || !cell.has_only_duration_format_metadata() {
        return Err(Error::UnsupportedDependency { path });
    }
    cell.set_duration_format_identifier_preserving_value(desired_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let output = cell
        .try_encode_with_limit(output_limit)
        .map_err(|error| match error {
            litchi_numbers_wire::Error::Allocation { requested } => Error::Allocation {
                amount: requested,
                path,
            },
            litchi_numbers_wire::Error::InvalidFormat(_)
            | litchi_numbers_wire::Error::ParseError(_)
            | litchi_numbers_wire::Error::OutputLimitExceeded { .. } => {
                Error::InvalidSource { path }
            },
        })?;
    let candidate_cell = BncCell::parse(&output).map_err(|_| Error::Verification)?;
    verify_duration_cell_metadata(
        source_value,
        source_cache,
        source_secondary,
        &candidate_cell,
        desired_identifier,
    )?;
    Ok(output)
}

fn verify_duration_cell_metadata(
    source_value: litchi_numbers_wire::StoredValue,
    source_cache: Option<litchi_numbers_wire::CachedScalar>,
    source_secondary: Option<u32>,
    candidate_cell: &BncCell,
    desired_identifier: Option<u32>,
) -> Result<(), Error> {
    let expected_secondary = desired_identifier.and(source_secondary);
    let expected_explicit = desired_identifier.map_or(0, |_| {
        litchi_numbers_wire::explicit_duration_format_flags(expected_secondary.is_some())
    });
    if source_value != candidate_cell.stored_value()
        || candidate_cell.cached_scalar().ok() != Some(source_cache)
        || !candidate_cell.is_duration_format_compatible()
        || !candidate_cell.has_only_duration_format_metadata()
    {
        return Err(Error::Verification);
    }
    if candidate_cell.explicit_format_flags() != expected_explicit
        || candidate_cell.cell_format_kind()
            != desired_identifier.map(|_| litchi_numbers_wire::DURATION_CELL_FORMAT_KIND)
        || candidate_cell.format_identifier() != desired_identifier
        || candidate_cell.secondary_format_identifier() != expected_secondary
        || candidate_cell.control_cell_spec_identifier().is_some()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn rewrite_currency_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_limit = source
        .len()
        .checked_add(12)
        .ok_or(Error::InvalidSource { path })?;
    let owned_cell_bytes = source
        .len()
        .checked_add(output_limit)
        .ok_or(Error::InvalidSource { path })?;
    let work = output_limit
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(source.len()))
        .ok_or(Error::InvalidSource { path })?;
    let allocations = litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS
        .checked_mul(2)
        // One encoded output plus the three small format-field buffers that
        // the Currency mutation can materialize.
        .and_then(|amount| amount.checked_add(4))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(allocations, path)?;
    budget.charge_scratch_bytes(owned_cell_bytes, path)?;
    budget.charge_retained_bytes(output_limit, path)?;
    budget.charge_transaction_work(work, path)?;

    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    let source_secondary = cell.secondary_format_identifier();
    cell.set_currency_format_identifier_preserving_value(desired_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let output = cell
        .try_encode_with_limit(output_limit)
        .map_err(|error| match error {
            litchi_numbers_wire::Error::Allocation { requested } => Error::Allocation {
                amount: requested,
                path,
            },
            litchi_numbers_wire::Error::InvalidFormat(_)
            | litchi_numbers_wire::Error::ParseError(_)
            | litchi_numbers_wire::Error::OutputLimitExceeded { .. } => {
                Error::InvalidSource { path }
            },
        })?;
    let candidate_cell = BncCell::parse(&output).map_err(|_| Error::Verification)?;
    verify_currency_cell_metadata(
        source_value,
        source_cache,
        source_secondary,
        &candidate_cell,
        desired_identifier,
    )?;
    Ok(output)
}

fn verify_currency_cell_metadata(
    source_value: litchi_numbers_wire::StoredValue,
    source_cache: Option<litchi_numbers_wire::CachedScalar>,
    source_secondary: Option<u32>,
    candidate_cell: &BncCell,
    desired_identifier: Option<u32>,
) -> Result<(), Error> {
    let expected_secondary = desired_identifier.and(source_secondary);
    let expected_explicit = desired_identifier.map_or(0, |_| {
        litchi_numbers_wire::explicit_currency_format_flags(expected_secondary.is_some())
    });
    if source_value != candidate_cell.stored_value()
        || candidate_cell.cached_scalar().ok() != Some(source_cache)
    {
        return Err(Error::Verification);
    }
    if candidate_cell.explicit_format_flags() != expected_explicit
        || !currency_cell_type_matches_value(candidate_cell, desired_identifier.is_some())
        || candidate_cell.cell_format_kind()
            != desired_identifier.map(|_| litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND)
        || candidate_cell.format_identifier() != desired_identifier
        || candidate_cell.secondary_format_identifier() != expected_secondary
        || candidate_cell.control_cell_spec_identifier().is_some()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn currency_cell_type_matches_value(cell: &BncCell, formatted: bool) -> bool {
    match cell.stored_value() {
        // Numbers can attach display metadata to an otherwise empty cell.
        // That shape has no numeric cell type and must remain empty through a
        // metadata-only set or clear operation.
        litchi_numbers_wire::StoredValue::Empty => cell.numeric_cell_type().is_none(),
        _ => {
            cell.numeric_cell_type()
                == Some(if formatted {
                    NumericCellType::AlternateNumber
                } else {
                    NumericCellType::Number
                })
        },
    }
}

fn scientific_cell_type_matches_value(cell: &BncCell) -> bool {
    match cell.stored_value() {
        litchi_numbers_wire::StoredValue::Empty => cell.numeric_cell_type().is_none(),
        litchi_numbers_wire::StoredValue::Number | litchi_numbers_wire::StoredValue::Formula(_) => {
            cell.numeric_cell_type() == Some(NumericCellType::Number)
        },
        _ => false,
    }
}

/// Reserve the private vectors used by the display-format tile patcher before it
/// enters the allocator.
///
/// `patch_tile_cell` is intentionally shared with the older control owners,
/// so it cannot borrow the operation ledger itself. Display-format edits still
/// need to account for its nested tile/row/buffer candidates before the first
/// `Vec` is materialized.  The bound includes the source tile, replacement,
/// and the three length-delimited framing layers; the scratch reservation is
/// deliberately conservative for the simultaneously-live row and buffer
/// candidates, while retained bytes cover the returned tile candidate.
fn charge_display_tile_patch_budget(
    tile_payload_len: usize,
    replacement_len: usize,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let candidate_bound = tile_payload_len
        .checked_add(replacement_len)
        .and_then(|length| length.checked_add(32))
        .ok_or(Error::InvalidSource { path })?;
    let scratch_bound = candidate_bound
        .checked_mul(3)
        .ok_or(Error::InvalidSource { path })?;
    let work_bound = candidate_bound
        .checked_add(scratch_bound)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(5, path)?;
    budget.charge_scratch_bytes(scratch_bound, path)?;
    budget.charge_retained_bytes(candidate_bound, path)?;
    budget.charge_transaction_work(work_bound, path)?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct TextCellMetadata {
    generic_identifier: Option<u32>,
    text_identifier: Option<u32>,
    text_identifier_offset: Option<usize>,
}

/// Inspect the fixed BNC metadata needed by the Text owner.
///
/// `BncCell` deliberately does not expose the generic secondary identifier
/// carried by native converted-text cells (`0x81`).  Keeping this small
/// parser private lets the package owner preserve that reference without
/// widening the low-level wire API or leaking an ID through the semantic
/// facade.  The preceding `BncCell::parse` still owns version, field-width,
/// and unknown-flag validation; these checks only recover offsets and reject
/// metadata from another family.
fn inspect_text_cell_metadata(
    source: &[u8],
    cell: &BncCell,
    path: Path,
) -> Result<TextCellMetadata, Error> {
    if source.len() < BNC_HEADER_LEN {
        return Err(Error::InvalidSource { path });
    }
    let flags = u32::from_le_bytes(
        source[BNC_HEADER_LEN - 4..BNC_HEADER_LEN]
            .try_into()
            .map_err(|_| Error::InvalidSource { path })?,
    );
    if flags & BNC_RESERVED_KNOWN_FIELD_FLAG != 0 {
        return Err(Error::UnsupportedDependency { path });
    }
    if flags & BNC_FORMAT_FLAGS & !BNC_TEXT_ALLOWED_FORMAT_FLAGS != 0 {
        return Err(Error::UnsupportedDependency { path });
    }
    if cell.cell_format_kind() != Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
        return Err(Error::UnsupportedDependency { path });
    }

    let mut offset = BNC_HEADER_LEN;
    let mut generic_identifier = None;
    let mut text_identifier = None;
    let mut text_identifier_offset = None;
    for &(flag, size) in BNC_FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        let end = offset
            .checked_add(size)
            .ok_or(Error::InvalidSource { path })?;
        let bytes = source
            .get(offset..end)
            .ok_or(Error::InvalidSource { path })?;
        if flag == BNC_CELL_FORMAT_IDENTIFIER_FLAG {
            generic_identifier = Some(u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_| Error::InvalidSource { path })?,
            ));
        } else if flag == BNC_TEXT_FORMAT_IDENTIFIER_FLAG {
            text_identifier = Some(u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_| Error::InvalidSource { path })?,
            ));
            text_identifier_offset = Some(offset);
        }
        offset = end;
    }
    if offset > source.len() {
        return Err(Error::InvalidSource { path });
    }
    Ok(TextCellMetadata {
        generic_identifier,
        text_identifier,
        text_identifier_offset,
    })
}

fn validate_text_cell_metadata(
    source: &[u8],
    cell: &BncCell,
    path: Path,
) -> Result<TextCellMetadata, Error> {
    let metadata = inspect_text_cell_metadata(source, cell, path)?;
    let explicit_flags = cell.explicit_format_flags();
    let valid_value = matches!(
        cell.stored_value(),
        litchi_numbers_wire::StoredValue::Empty | litchi_numbers_wire::StoredValue::Text(_)
    );
    let valid_marker = match explicit_flags {
        // A native Text cell may carry its Text key with no explicit marker
        // when it is still using the workbook's automatic/default display.
        // Treat that shape as a valid Text owner, but never publish it when a
        // caller explicitly stages Text (the write path uses 0x80).
        0 | litchi_numbers_wire::EXPLICIT_TEXT_FORMAT => metadata.generic_identifier.is_none(),
        litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT => metadata
            .generic_identifier
            .is_some_and(|identifier| identifier != 0),
        _ => false,
    };
    if !valid_value
        || !valid_marker
        || cell.control_cell_spec_identifier().is_some()
        || metadata
            .text_identifier
            .is_none_or(|identifier| identifier == 0)
    {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(metadata)
}

fn rewrite_text_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_limit = source
        .len()
        .checked_add(8)
        .ok_or(Error::InvalidSource { path })?;
    let owned_cell_bytes = source
        .len()
        .checked_add(output_limit)
        .ok_or(Error::InvalidSource { path })?;
    let work = output_limit
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(source.len()))
        .ok_or(Error::InvalidSource { path })?;
    let allocations = litchi_numbers_wire::MAX_OWNED_BNC_PARSE_ALLOCATIONS
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(2))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(allocations, path)?;
    budget.charge_scratch_bytes(owned_cell_bytes, path)?;
    budget.charge_retained_bytes(output_limit, path)?;
    budget.charge_transaction_work(work, path)?;

    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    let metadata = if cell.cell_format_kind() == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
        Some(validate_text_cell_metadata(source, &cell, path)?)
    } else {
        None
    };

    // Installing a Text display format on an automatic numeric/formula/date
    // cell would make the generic BNC setter convert its scalar. This owner
    // is metadata-only: only an empty or already-text value may enter the
    // canonical Text path.
    if desired_identifier.is_some()
        && metadata.is_none()
        && !matches!(
            source_value,
            litchi_numbers_wire::StoredValue::Empty | litchi_numbers_wire::StoredValue::Text(_)
        )
    {
        return Err(Error::UnsupportedDependency { path });
    }

    let output = match desired_identifier {
        Some(0) => return Err(Error::InvalidSource { path }),
        Some(identifier) => {
            if let Some(metadata) = metadata {
                // Existing native Text cells have a fixed-width text key. A
                // direct byte patch is the only way to retain the converted
                // Text marker and its generic Number reference exactly. The
                // marker-zero/default shape is promoted to plain explicit
                // Text, never to converted Text.
                let offset = metadata
                    .text_identifier_offset
                    .ok_or(Error::InvalidSource { path })?;
                let mut output = Vec::new();
                output
                    .try_reserve_exact(source.len())
                    .map_err(|_| Error::Allocation {
                        amount: source.len(),
                        path,
                    })?;
                output.extend_from_slice(source);
                output[offset..offset + 4].copy_from_slice(&identifier.to_le_bytes());
                if cell.explicit_format_flags() == 0 {
                    output[BNC_EXPLICIT_FORMAT_FLAGS_START..BNC_EXPLICIT_FORMAT_FLAGS_END]
                        .copy_from_slice(&litchi_numbers_wire::EXPLICIT_TEXT_FORMAT.to_le_bytes());
                }
                output
            } else {
                cell.set_data_format_identifier(identifier, CellDataFormatKind::Text, None)
                    .map_err(|_| Error::UnsupportedDependency { path })?;
                cell.try_encode_with_limit(output_limit)
                    .map_err(|error| map_bnc_error(error, path))?
            }
        },
        None => {
            cell.clear_explicit_format();
            cell.try_encode_with_limit(output_limit)
                .map_err(|error| map_bnc_error(error, path))?
        },
    };
    let candidate = BncCell::parse(&output).map_err(|_| Error::Verification)?;
    if candidate.stored_value() != source_value
        || candidate.cached_scalar().map_err(|_| Error::Verification)? != source_cache
    {
        return Err(Error::Verification);
    }
    match desired_identifier {
        Some(identifier) => {
            let candidate_metadata = validate_text_cell_metadata(&output, &candidate, path)?;
            if candidate_metadata.text_identifier != Some(identifier)
                || metadata.is_some_and(|source_metadata| {
                    candidate_metadata.generic_identifier != source_metadata.generic_identifier
                })
            {
                return Err(Error::Verification);
            }
        },
        None => {
            if candidate.explicit_format_flags() != 0
                || candidate.cell_format_kind().is_some()
                || candidate.format_identifier().is_some()
                || candidate.control_cell_spec_identifier().is_some()
            {
                return Err(Error::Verification);
            }
        },
    }
    Ok(output)
}

fn map_bnc_error(error: litchi_numbers_wire::Error, path: Path) -> Error {
    match error {
        litchi_numbers_wire::Error::Allocation { requested } => Error::Allocation {
            amount: requested,
            path,
        },
        litchi_numbers_wire::Error::InvalidFormat(_)
        | litchi_numbers_wire::Error::ParseError(_)
        | litchi_numbers_wire::Error::OutputLimitExceeded { .. } => Error::InvalidSource { path },
    }
}

/// Read one existing Number format with a fresh transaction ledger.
pub(super) fn read_number_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Number>, NumberFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_number_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Number format against a caller-owned transaction ledger.
pub(super) fn read_number_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Number>, NumberFormatReadError> {
    match read_display_format_with_budget(DisplayFormatFamily::Number, source, target, path, budget)
    {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Number(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => Err(NumberFormatReadError::WrongFormatFamily),
        Err(DisplayReadError::Native(error)) => Err(NumberFormatReadError::Native(error)),
    }
}

/// Read one existing Currency format with a fresh transaction ledger.
pub(super) fn read_currency_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Currency>, CurrencyFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_currency_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Currency format against a caller-owned transaction
/// ledger.
pub(super) fn read_currency_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Currency>, CurrencyFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::Currency,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Currency(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(CurrencyFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(CurrencyFormatReadError::Native(error)),
    }
}

/// Read one existing Percentage format with a fresh transaction ledger.
pub(super) fn read_percentage_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Percentage>, PercentageFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_percentage_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Percentage format against a caller-owned transaction
/// ledger.
pub(super) fn read_percentage_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Percentage>, PercentageFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::Percentage,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Percentage(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(PercentageFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(PercentageFormatReadError::Native(error)),
    }
}

/// Read one existing Scientific format with a fresh transaction ledger.
pub(super) fn read_scientific_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Scientific>, ScientificFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_scientific_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Scientific format against a caller-owned transaction
/// ledger.
pub(super) fn read_scientific_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Scientific>, ScientificFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::Scientific,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Scientific(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(ScientificFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(ScientificFormatReadError::Native(error)),
    }
}

/// Read one existing Fraction format with a fresh transaction ledger.
pub(super) fn read_fraction_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Fraction>, FractionFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_fraction_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Fraction format against a caller-owned transaction
/// ledger.
pub(super) fn read_fraction_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Fraction>, FractionFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::Fraction,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Fraction(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(FractionFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(FractionFormatReadError::Native(error)),
    }
}

/// Read one existing Date & Time format with a fresh transaction ledger.
pub(super) fn read_date_time_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<DateTime>, DateTimeFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_date_time_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Date & Time format against a caller-owned transaction
/// ledger.
pub(super) fn read_date_time_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<DateTime>, DateTimeFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::DateTime,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::DateTime(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(DateTimeFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(DateTimeFormatReadError::Native(error)),
    }
}

/// Read one existing Duration format with a fresh transaction ledger.
pub(super) fn read_duration_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Duration>, DurationFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_duration_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Duration format against a caller-owned transaction
/// ledger.
pub(super) fn read_duration_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Duration>, DurationFormatReadError> {
    match read_display_format_with_budget(
        DisplayFormatFamily::Duration,
        source,
        target,
        path,
        budget,
    ) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Duration(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Text(_)))
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(DurationFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(DurationFormatReadError::Native(error)),
    }
}

/// Read one existing Text format with a fresh transaction ledger.
pub(super) fn read_text_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Text>, TextFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_text_format_with_budget(source, target, path, &mut budget)
}

/// Read one existing Text format against a caller-owned transaction ledger.
pub(super) fn read_text_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Text>, TextFormatReadError> {
    match read_display_format_with_budget(DisplayFormatFamily::Text, source, target, path, budget) {
        Ok(None) => Ok(None),
        Ok(Some(NativeDisplayValue::Text(value))) => Ok(Some(value)),
        Ok(Some(NativeDisplayValue::Number(_)))
        | Ok(Some(NativeDisplayValue::Currency(_)))
        | Ok(Some(NativeDisplayValue::Percentage(_)))
        | Ok(Some(NativeDisplayValue::Scientific(_)))
        | Ok(Some(NativeDisplayValue::Fraction(_)))
        | Ok(Some(NativeDisplayValue::DateTime(_)))
        | Ok(Some(NativeDisplayValue::Duration(_)))
        | Err(DisplayReadError::WrongFormatFamily) => Err(TextFormatReadError::WrongFormatFamily),
        Err(DisplayReadError::Native(error)) => Err(TextFormatReadError::Native(error)),
    }
}

/// Rewrite one ordinary Number cell without manufacturing a CellSpec graph.
pub(super) fn rewrite_number_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Number>,
    after: Option<&Number>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Number,
        source,
        target,
        before.copied().map(NativeDisplayValue::Number),
        after.copied().map(NativeDisplayValue::Number),
        path,
        budget,
    )
}

/// Rewrite one ordinary Currency cell without manufacturing a CellSpec graph.
pub(super) fn rewrite_currency_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Currency>,
    after: Option<&Currency>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Currency,
        source,
        target,
        before.copied().map(NativeDisplayValue::Currency),
        after.copied().map(NativeDisplayValue::Currency),
        path,
        budget,
    )
}

/// Rewrite one ordinary Percentage cell without manufacturing a CellSpec
/// graph.
pub(super) fn rewrite_percentage_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Percentage>,
    after: Option<&Percentage>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Percentage,
        source,
        target,
        before.copied().map(NativeDisplayValue::Percentage),
        after.copied().map(NativeDisplayValue::Percentage),
        path,
        budget,
    )
}

/// Rewrite one ordinary Scientific cell without manufacturing a CellSpec
/// graph.
pub(super) fn rewrite_scientific_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Scientific>,
    after: Option<&Scientific>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Scientific,
        source,
        target,
        before.copied().map(NativeDisplayValue::Scientific),
        after.copied().map(NativeDisplayValue::Scientific),
        path,
        budget,
    )
}

/// Rewrite one ordinary Fraction cell without manufacturing a CellSpec graph.
pub(super) fn rewrite_fraction_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Fraction>,
    after: Option<&Fraction>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Fraction,
        source,
        target,
        before.copied().map(NativeDisplayValue::Fraction),
        after.copied().map(NativeDisplayValue::Fraction),
        path,
        budget,
    )
}

/// Rewrite one Date & Time cell without converting its stored value or
/// formula cache.
pub(super) fn rewrite_date_time_format(
    source: &Package,
    target: CellTarget,
    before: Option<&DateTime>,
    after: Option<&DateTime>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::DateTime,
        source,
        target,
        before.cloned().map(NativeDisplayValue::DateTime),
        after.cloned().map(NativeDisplayValue::DateTime),
        path,
        budget,
    )
}

/// Rewrite one ordinary Duration cell without changing its stored value or
/// formula cache.
pub(super) fn rewrite_duration_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Duration>,
    after: Option<&Duration>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Duration,
        source,
        target,
        before.copied().map(NativeDisplayValue::Duration),
        after.copied().map(NativeDisplayValue::Duration),
        path,
        budget,
    )
}

/// Rewrite one ordinary Text cell without manufacturing a CellSpec graph.
pub(super) fn rewrite_text_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Text>,
    after: Option<&Text>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    rewrite_display_format(
        DisplayFormatFamily::Text,
        source,
        target,
        before.copied().map(NativeDisplayValue::Text),
        after.copied().map(NativeDisplayValue::Text),
        path,
        budget,
    )
}

fn read_display_format_with_budget(
    family: DisplayFormatFamily,
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<NativeDisplayValue>, DisplayReadError> {
    let model_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let model_archive = model_component.archive();
    let model_object = native::unique_object(model_archive, target.model_identifier, path)?;
    let model_index = native::unique_message_index(model_object, 6_001, path)?;
    let model_payload = &model_object.messages[model_index].data;
    let (model, model_report) = storage_codec::decode_table_model_with_report(
        model_payload,
        budget.residual_storage_options(model_payload),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, model_report, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model.base_data_store(),
        budget.residual_storage_options(model.base_data_store()),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, store_report, path)?;
    let tile_options = budget.residual_storage_options(store.tiles());
    let mut tile_visitor = native::TileReferenceCollector::new(budget, path);
    let decoded_tiles = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        tile_options,
        &mut tile_visitor,
    );
    let tile_references = tile_visitor.finish()?;
    let (tile_storage, tile_report) =
        decoded_tiles.map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, tile_report, path)?;
    let tile_size = tile_storage
        .tile_size()
        .filter(|size| *size != 0)
        .ok_or(Error::InvalidSource { path })?;
    if tile_references.iter().any(|(_, reference)| *reference == 0)
        || tile_references
            .iter()
            .enumerate()
            .any(|(index, (id, reference))| {
                tile_references[index + 1..]
                    .iter()
                    .any(|(other_id, other_reference)| {
                        id == other_id || reference == other_reference
                    })
            })
    {
        return Err(Error::InvalidSource { path }.into());
    }
    let tile_id = target.position.row() / tile_size;
    let mut tile_matches = tile_references
        .iter()
        .filter(|(id, _)| *id == tile_id)
        .map(|(_, reference)| *reference);
    let tile_identifier = tile_matches.next().ok_or(Error::CellNotFound)?;
    if tile_matches.next().is_some() || tile_identifier == target.model_identifier {
        return Err(Error::InvalidSource { path }.into());
    }
    popup::validate_selected_model_reference(source, target, tile_identifier, path)?;
    let tile_component_index = native::resolved_component_index(source, tile_identifier, path)?;
    let tile_archive = source
        .state
        .components
        .catalog()
        .get_index(tile_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let tile_object = native::unique_object(tile_archive, tile_identifier, path)?;
    let tile_message_index = native::unique_message_index(tile_object, 6_002, path)?;
    let tile_payload = &tile_object.messages[tile_message_index].data;
    let cell_source = popup_native::tile_cell(
        tile_payload,
        target.position.row(),
        target.position.column(),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    charge_owned_bnc_parse(cell_source.len(), path, budget)?;
    let cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
    native::validate_bnc_format_metadata(cell_source, &cell, path)?;
    let format_identifier = cell.format_identifier();
    if cell.control_cell_spec_identifier().is_some() {
        return Err(DisplayReadError::WrongFormatFamily);
    }
    let explicit_flags = cell.explicit_format_flags();
    let secondary_identifier =
        if cell.cell_format_kind() == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
            Some(validate_text_cell_metadata(cell_source, &cell, path)?)
                .and_then(|metadata| metadata.generic_identifier)
        } else {
            cell.secondary_format_identifier()
        };
    match (family, cell.cell_format_kind()) {
        (
            DisplayFormatFamily::Number
            | DisplayFormatFamily::Percentage
            | DisplayFormatFamily::Scientific
            | DisplayFormatFamily::Fraction,
            Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND),
        ) => {
            if secondary_identifier.is_some() || !cell.has_only_decimal_format_metadata() {
                return Err(DisplayReadError::WrongFormatFamily);
            }
            if explicit_flags != 0 && explicit_flags != litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT
            {
                return Err(DisplayReadError::WrongFormatFamily);
            }
            if family == DisplayFormatFamily::Scientific
                && !scientific_cell_type_matches_value(&cell)
            {
                return Err(Error::InvalidSource { path }.into());
            }
        },
        (DisplayFormatFamily::DateTime, Some(litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND)) => {
            if secondary_identifier.is_some()
                || !cell.has_only_date_time_format_metadata()
                || explicit_flags != litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT
                || !cell.is_date_time_format_compatible()
            {
                return Err(DisplayReadError::WrongFormatFamily);
            }
        },
        (DisplayFormatFamily::Duration, Some(litchi_numbers_wire::DURATION_CELL_FORMAT_KIND)) => {
            let expected_explicit =
                litchi_numbers_wire::explicit_duration_format_flags(secondary_identifier.is_some());
            if explicit_flags != expected_explicit
                || secondary_identifier.is_some_and(|identifier| identifier == 0)
                || !cell.has_only_duration_format_metadata()
                || !cell.is_duration_format_compatible()
            {
                return Err(DisplayReadError::WrongFormatFamily);
            }
        },
        (DisplayFormatFamily::Currency, Some(litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND)) => {
            let expected_explicit =
                litchi_numbers_wire::explicit_currency_format_flags(secondary_identifier.is_some());
            if explicit_flags != 0 && explicit_flags != expected_explicit
                || secondary_identifier.is_some_and(|identifier| identifier == 0)
                || !currency_cell_type_matches_value(&cell, true)
            {
                return Err(Error::InvalidSource { path }.into());
            }
        },
        (DisplayFormatFamily::Text, Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND)) => {
            let valid_marker = match explicit_flags {
                0 | litchi_numbers_wire::EXPLICIT_TEXT_FORMAT => secondary_identifier.is_none(),
                litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT => {
                    secondary_identifier.is_some_and(|identifier| identifier != 0)
                },
                _ => false,
            };
            if !valid_marker
                || !matches!(
                    cell.stored_value(),
                    litchi_numbers_wire::StoredValue::Empty
                        | litchi_numbers_wire::StoredValue::Text(_)
                )
            {
                return Err(DisplayReadError::WrongFormatFamily);
            }
        },
        (_, Some(_)) => return Err(DisplayReadError::WrongFormatFamily),
        (_, None) if format_identifier.is_some() || explicit_flags != 0 => {
            return Err(DisplayReadError::WrongFormatFamily);
        },
        (DisplayFormatFamily::DateTime, None) if !cell.is_date_time_format_compatible() => {
            return Err(DisplayReadError::WrongFormatFamily);
        },
        (DisplayFormatFamily::Duration, None) if !cell.is_duration_format_compatible() => {
            return Err(DisplayReadError::WrongFormatFamily);
        },
        (_, None) => return Ok(None),
    }
    let format_identifier = format_identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource { path })?;

    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    if format_table_identifier == target.model_identifier
        || format_table_identifier == tile_identifier
    {
        return Err(Error::UnsupportedDependency { path }.into());
    }
    popup::validate_selected_model_reference(source, target, format_table_identifier, path)?;
    native::require_exclusive_selected_model_reference(
        source,
        target,
        format_table_identifier,
        path,
        budget,
    )?;
    let format_component_index =
        native::resolved_component_index(source, format_table_identifier, path)?;
    let format_archive = source
        .state
        .components
        .catalog()
        .get_index(format_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let metadata_facts = native::validate_metadata_ownership(
        source,
        &[
            target.component_index,
            tile_component_index,
            format_component_index,
        ],
        budget,
        path,
    )?;
    native::validate_cross_component_reference(
        &metadata_facts,
        target.component_index,
        tile_component_index,
        tile_identifier,
        path,
    )?;
    native::validate_cross_component_reference(
        &metadata_facts,
        target.component_index,
        format_component_index,
        format_table_identifier,
        path,
    )?;
    metadata_facts
        .current_uuids_if_registered(&[
            (target.component_index, target.model_identifier),
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
        ])
        .map_err(|_| Error::UnsupportedDependency { path })?;
    let format_object = native::unique_object(format_archive, format_table_identifier, path)?;
    let format_message = native::unique_list_message(format_object, 2, budget, path)?;
    let format_payload = native::copy_payload_with_budget(format_message.payload, budget, path)?;
    let format_facts = native::list_facts(&format_payload, budget, path)?;
    if format_facts
        .entries
        .iter()
        .any(|entry| !entry.is_format || entry.payload.is_empty())
    {
        return Err(Error::InvalidSource { path }.into());
    }
    native::validate_format_refcounts(
        source,
        &tile_references,
        format_facts.entries.as_slice(),
        budget,
        path,
    )?;
    let entry = format_facts
        .entries
        .iter()
        .find(|entry| entry.key == format_identifier)
        .ok_or(Error::InvalidSource { path })?;
    if entry.ref_count == 0 || !entry.is_format {
        return Err(Error::InvalidSource { path }.into());
    }
    if let Some(secondary_identifier) = secondary_identifier {
        let secondary_entry = format_facts
            .entries
            .iter()
            .find(|entry| entry.key == secondary_identifier)
            .ok_or(Error::InvalidSource { path })?;
        if secondary_entry.ref_count == 0 || !secondary_entry.is_format {
            return Err(Error::InvalidSource { path }.into());
        }
        let (secondary, report) = number_codec::decode_number_format_with_report(
            &secondary_entry.payload,
            native::control_codec_options(secondary_entry.payload.len(), budget),
        )
        .map_err(|error| native::map_control_error(error, path))?;
        native::charge_control_decode_report(budget, report, path)?;
        if secondary.format_type() != NUMBER_FORMAT_TYPE {
            return Err(Error::InvalidSource { path }.into());
        }
    }
    let value = decode_display_payload(family, &entry.payload, budget, path)?;
    if explicit_flags == 0 {
        Ok(None)
    } else {
        Ok(Some(value))
    }
}

fn rewrite_display_format(
    family: DisplayFormatFamily,
    source: &Package,
    target: CellTarget,
    before: Option<NativeDisplayValue>,
    after: Option<NativeDisplayValue>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    if target.locked {
        return Err(Error::TableLocked { path });
    }
    if before
        .as_ref()
        .is_some_and(|value| value.family() != family)
        || after.as_ref().is_some_and(|value| value.family() != family)
    {
        return Err(Error::UnsupportedDependency { path });
    }
    let model_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let model_archive = model_component.archive();
    let model_object = native::unique_object(model_archive, target.model_identifier, path)?;
    let model_index = native::unique_message_index(model_object, 6_001, path)?;
    let model_payload = &model_object.messages[model_index].data;
    let (model, model_report) = storage_codec::decode_table_model_with_report(
        model_payload,
        budget.residual_storage_options(model_payload),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, model_report, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model.base_data_store(),
        budget.residual_storage_options(model.base_data_store()),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, store_report, path)?;
    let tile_options = budget.residual_storage_options(store.tiles());
    let mut tile_visitor = native::TileReferenceCollector::new(budget, path);
    let decoded_tiles = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        tile_options,
        &mut tile_visitor,
    );
    let tile_references = tile_visitor.finish()?;
    let (tile_storage, tile_report) =
        decoded_tiles.map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, tile_report, path)?;
    let tile_size = tile_storage
        .tile_size()
        .filter(|size| *size != 0)
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target.position.row() / tile_size;
    let mut tile_matches = tile_references
        .iter()
        .filter(|(id, _)| *id == tile_id)
        .map(|(_, reference)| *reference);
    let tile_identifier = tile_matches.next().ok_or(Error::CellNotFound)?;
    if tile_matches.next().is_some() || tile_identifier == target.model_identifier {
        return Err(Error::InvalidSource { path });
    }
    popup::validate_selected_model_reference(source, target, tile_identifier, path)?;
    native::require_exclusive_selected_model_reference(
        source,
        target,
        tile_identifier,
        path,
        budget,
    )?;
    if tile_references.iter().any(|(_, reference)| *reference == 0)
        || tile_references
            .iter()
            .enumerate()
            .any(|(index, (id, reference))| {
                tile_references[index + 1..]
                    .iter()
                    .any(|(other_id, other_reference)| {
                        id == other_id || reference == other_reference
                    })
            })
    {
        return Err(Error::InvalidSource { path });
    }
    let tile_component_index = native::resolved_component_index(source, tile_identifier, path)?;
    let tile_archive = source
        .state
        .components
        .catalog()
        .get_index(tile_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let tile_object = native::unique_object(tile_archive, tile_identifier, path)?;
    let tile_message_index = native::unique_message_index(tile_object, 6_002, path)?;
    let tile_payload = native::copy_payload_with_budget(
        &tile_object.messages[tile_message_index].data,
        budget,
        path,
    )?;
    let cell_source = popup_native::tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    charge_owned_bnc_parse(cell_source.len(), path, budget)?;
    let cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
    native::validate_bnc_format_metadata(cell_source, &cell, path)?;
    let old_format = cell.format_identifier();
    let old_secondary =
        if cell.cell_format_kind() == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
            Some(validate_text_cell_metadata(cell_source, &cell, path)?)
                .and_then(|metadata| metadata.generic_identifier)
        } else {
            cell.secondary_format_identifier()
        };
    if cell.control_cell_spec_identifier().is_some() {
        return Err(Error::UnsupportedDependency { path });
    }
    let explicit_flags = cell.explicit_format_flags();
    match (family, old_format, cell.cell_format_kind()) {
        (
            DisplayFormatFamily::Number
            | DisplayFormatFamily::Percentage
            | DisplayFormatFamily::Scientific
            | DisplayFormatFamily::Fraction,
            Some(identifier),
            Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND),
        ) if identifier != 0 => {
            if old_secondary.is_some()
                || !cell.has_only_decimal_format_metadata()
                || (explicit_flags != 0
                    && explicit_flags != litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT)
                || (family == DisplayFormatFamily::Scientific
                    && !scientific_cell_type_matches_value(&cell))
            {
                return Err(Error::UnsupportedDependency { path });
            }
        },
        (DisplayFormatFamily::DateTime, Some(identifier), Some(kind))
            if identifier != 0 && kind == litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND =>
        {
            if old_secondary.is_some()
                || !cell.has_only_date_time_format_metadata()
                || explicit_flags != litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT
                || !cell.is_date_time_format_compatible()
            {
                return Err(Error::UnsupportedDependency { path });
            }
        },
        (DisplayFormatFamily::Duration, Some(identifier), Some(kind))
            if identifier != 0 && kind == litchi_numbers_wire::DURATION_CELL_FORMAT_KIND =>
        {
            let expected_explicit =
                litchi_numbers_wire::explicit_duration_format_flags(old_secondary.is_some());
            if old_secondary.is_some_and(|identifier| identifier == 0)
                || !cell.has_only_duration_format_metadata()
                || explicit_flags != expected_explicit
                || !cell.is_duration_format_compatible()
            {
                return Err(Error::UnsupportedDependency { path });
            }
        },
        (DisplayFormatFamily::Currency, Some(identifier), Some(kind))
            if identifier != 0 && kind == litchi_numbers_wire::CURRENCY_CELL_FORMAT_KIND =>
        {
            let expected_explicit =
                litchi_numbers_wire::explicit_currency_format_flags(old_secondary.is_some());
            if explicit_flags != 0 && explicit_flags != expected_explicit
                || old_secondary.is_some_and(|identifier| identifier == 0)
                || !currency_cell_type_matches_value(&cell, true)
            {
                return Err(Error::UnsupportedDependency { path });
            }
        },
        (DisplayFormatFamily::Text, Some(identifier), Some(kind))
            if identifier != 0 && kind == litchi_numbers_wire::TEXT_CELL_FORMAT_KIND =>
        {
            let metadata = validate_text_cell_metadata(cell_source, &cell, path)?;
            let expected_marker = match explicit_flags {
                0 => old_secondary.is_none() && before.is_none(),
                litchi_numbers_wire::EXPLICIT_TEXT_FORMAT => old_secondary.is_none(),
                litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT => {
                    old_secondary.is_some_and(|identifier| identifier != 0)
                },
                _ => false,
            };
            if !expected_marker
                || metadata.text_identifier != Some(identifier)
                || old_secondary != metadata.generic_identifier
            {
                return Err(Error::UnsupportedDependency { path });
            }
        },
        (_, None, None)
            if explicit_flags == 0
                && old_secondary.is_none()
                && (family != DisplayFormatFamily::Currency
                    || currency_cell_type_matches_value(&cell, false))
                && (family != DisplayFormatFamily::Text
                    || matches!(
                        cell.stored_value(),
                        litchi_numbers_wire::StoredValue::Empty
                            | litchi_numbers_wire::StoredValue::Text(_)
                    ))
                && (family != DisplayFormatFamily::DateTime
                    || cell.is_date_time_format_compatible())
                && (family != DisplayFormatFamily::Duration
                    || cell.is_duration_format_compatible()) => {},
        _ => return Err(Error::UnsupportedDependency { path }),
    }
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    if format_table_identifier == target.model_identifier
        || format_table_identifier == tile_identifier
    {
        return Err(Error::UnsupportedDependency { path });
    }
    popup::validate_selected_model_reference(source, target, format_table_identifier, path)?;
    native::require_exclusive_selected_model_reference(
        source,
        target,
        format_table_identifier,
        path,
        budget,
    )?;
    let format_component_index =
        native::resolved_component_index(source, format_table_identifier, path)?;
    let format_archive = source
        .state
        .components
        .catalog()
        .get_index(format_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let metadata_facts = native::validate_metadata_ownership(
        source,
        &[
            target.component_index,
            tile_component_index,
            format_component_index,
        ],
        budget,
        path,
    )?;
    native::validate_cross_component_reference(
        &metadata_facts,
        target.component_index,
        tile_component_index,
        tile_identifier,
        path,
    )?;
    native::validate_cross_component_reference(
        &metadata_facts,
        target.component_index,
        format_component_index,
        format_table_identifier,
        path,
    )?;
    metadata_facts
        .current_uuids_if_registered(&[
            (target.component_index, target.model_identifier),
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
        ])
        .map_err(|_| Error::UnsupportedDependency { path })?;
    let format_object = native::unique_object(format_archive, format_table_identifier, path)?;
    let format_message = native::unique_list_message(format_object, 2, budget, path)?;
    let format_message_index = format_message.message_index;
    let format_payload = native::copy_payload_with_budget(format_message.payload, budget, path)?;
    let format_facts = native::list_facts(&format_payload, budget, path)?;
    if format_facts
        .entries
        .iter()
        .any(|entry| !entry.is_format || entry.payload.is_empty())
    {
        return Err(Error::InvalidSource { path });
    }
    native::validate_format_refcounts(
        source,
        &tile_references,
        format_facts.entries.as_slice(),
        budget,
        path,
    )?;

    if matches!(
        family,
        DisplayFormatFamily::Currency | DisplayFormatFamily::Duration | DisplayFormatFamily::Text
    ) {
        if let Some(secondary_identifier) = old_secondary {
            let secondary_entry = format_facts
                .entries
                .iter()
                .find(|entry| entry.key == secondary_identifier)
                .ok_or(Error::InvalidSource { path })?;
            if secondary_entry.ref_count == 0 || !secondary_entry.is_format {
                return Err(Error::InvalidSource { path });
            }
            let (secondary, report) = number_codec::decode_number_format_with_report(
                &secondary_entry.payload,
                native::control_codec_options(secondary_entry.payload.len(), budget),
            )
            .map_err(|error| native::map_control_error(error, path))?;
            native::charge_control_decode_report(budget, report, path)?;
            if secondary.format_type() != NUMBER_FORMAT_TYPE {
                return Err(Error::InvalidSource { path });
            }
        }
    }

    if let Some(old_key) = old_format {
        let entry = format_facts
            .entries
            .iter()
            .find(|entry| entry.key == old_key)
            .ok_or(Error::InvalidSource { path })?;
        if entry.ref_count == 0 || !entry.is_format {
            return Err(Error::InvalidSource { path });
        }
        let current = decode_display_payload(family, &entry.payload, budget, path)
            .map_err(|error| display_read_error_to_write_error(error, path))?;
        let expected_explicit = match family {
            DisplayFormatFamily::Currency => {
                litchi_numbers_wire::explicit_currency_format_flags(old_secondary.is_some())
            },
            DisplayFormatFamily::Number
            | DisplayFormatFamily::Percentage
            | DisplayFormatFamily::Scientific
            | DisplayFormatFamily::Fraction => litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT,
            DisplayFormatFamily::DateTime => litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT,
            DisplayFormatFamily::Duration => {
                litchi_numbers_wire::explicit_duration_format_flags(old_secondary.is_some())
            },
            DisplayFormatFamily::Text => {
                if old_secondary.is_some() {
                    litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT
                } else {
                    litchi_numbers_wire::EXPLICIT_TEXT_FORMAT
                }
            },
        };
        if explicit_flags == expected_explicit {
            if before != Some(current) {
                return Err(Error::PatchConflict);
            }
        } else if before.is_some() {
            return Err(Error::PatchConflict);
        }
    } else if before.is_some() {
        return Err(Error::PatchConflict);
    }

    let after_is_some = after.is_some();
    let after_is_none = !after_is_some;
    let desired_payload = match (after, old_format) {
        (Some(value), Some(old_key)) => {
            let entry = format_facts
                .entries
                .iter()
                .find(|entry| entry.key == old_key)
                .ok_or(Error::InvalidSource { path })?;
            Some(prepare_display_rewrite(
                family,
                &entry.payload,
                value,
                budget,
                path,
            )?)
        },
        (Some(value), None) => Some(prepare_display_append(
            family,
            value,
            format_payload.len(),
            budget,
            path,
        )?),
        (None, _) => None,
    };
    let (mut new_format, new_format_key) = native::mutate_list(
        &format_payload,
        old_format,
        desired_payload.as_deref(),
        true,
        budget,
        path,
    )?;
    if matches!(
        family,
        DisplayFormatFamily::Currency | DisplayFormatFamily::Duration | DisplayFormatFamily::Text
    ) && after_is_none
    {
        if let Some(secondary_identifier) = old_secondary {
            (new_format, _) = native::mutate_list(
                &new_format,
                Some(secondary_identifier),
                None,
                true,
                budget,
                path,
            )?;
        }
    }
    let replacement_cell = if after_is_some {
        let key = new_format_key.ok_or(Error::InvalidSource { path })?;
        if family == DisplayFormatFamily::Currency {
            rewrite_currency_cell_metadata(cell_source, Some(key), path, budget)?
        } else if family == DisplayFormatFamily::DateTime {
            rewrite_date_time_cell_metadata(cell_source, Some(key), path, budget)?
        } else if family == DisplayFormatFamily::Duration {
            rewrite_duration_cell_metadata(cell_source, Some(key), path, budget)?
        } else if family == DisplayFormatFamily::Text {
            rewrite_text_cell_metadata(cell_source, Some(key), path, budget)?
        } else {
            rewrite_display_cell_metadata(cell_source, Some(key), path, budget)?
        }
    } else {
        if family == DisplayFormatFamily::Currency {
            rewrite_currency_cell_metadata(cell_source, None, path, budget)?
        } else if family == DisplayFormatFamily::DateTime {
            rewrite_date_time_cell_metadata(cell_source, None, path, budget)?
        } else if family == DisplayFormatFamily::Duration {
            rewrite_duration_cell_metadata(cell_source, None, path, budget)?
        } else if family == DisplayFormatFamily::Text {
            rewrite_text_cell_metadata(cell_source, None, path, budget)?
        } else {
            rewrite_display_cell_metadata(cell_source, None, path, budget)?
        }
    };
    charge_display_tile_patch_budget(tile_payload.len(), replacement_cell.len(), budget, path)?;
    let patched_tile = popup_native::patch_tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
        &replacement_cell,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let mut component_indices = vec![tile_component_index, format_component_index];
    component_indices.sort_unstable();
    component_indices.dedup();
    let mut archives = component_indices
        .iter()
        .map(|&component_index| {
            let component = source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?;
            budget.charge_archive(
                component.archive(),
                native::archive_serialized_bound(component.archive(), archive_limits, path)?,
                path,
            )?;
            Ok::<_, Error>((component_index, component.archive().clone()))
        })
        .collect::<Result<Vec<_>, Error>>()?;
    popup_native::validate_message_without_object_references(
        format_archive,
        format_table_identifier,
        format_message_index,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        native::archive_for_mut(&mut archives, format_component_index, path)?,
        format_table_identifier,
        format_message_index,
        new_format,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        native::archive_for_mut(&mut archives, tile_component_index, path)?,
        tile_identifier,
        tile_message_index,
        patched_tile,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;

    let mut source_serialized_bytes = 0usize;
    let mut candidate_serialized_bytes = 0usize;
    for (component_index, candidate) in &archives {
        let source_archive = source
            .state
            .components
            .catalog()
            .get_index(*component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive();
        let source_length = source_archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let candidate_length = candidate
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        source_serialized_bytes = source_serialized_bytes
            .checked_add(source_length)
            .ok_or(Error::InvalidSource { path })?;
        candidate_serialized_bytes = candidate_serialized_bytes
            .checked_add(candidate_length)
            .ok_or(Error::InvalidSource { path })?;
    }
    let comparison_bytes = source_serialized_bytes
        .checked_add(candidate_serialized_bytes)
        .ok_or(Error::InvalidSource { path })?;
    let serialization_allocations = archives
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(serialization_allocations, path)?;
    budget.charge_scratch_bytes(comparison_bytes, path)?;
    budget.charge_retained_bytes(candidate_serialized_bytes, path)?;
    budget.charge_transaction_work(comparison_bytes, path)?;

    let mut members = Vec::new();
    members
        .try_reserve_exact(archives.len())
        .map_err(|_| Error::Allocation {
            amount: archives.len(),
            path,
        })?;
    for (component_index, archive) in archives {
        let changed_identifiers = [
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
        ]
        .iter()
        .filter_map(|(owner, identifier)| (*owner == component_index).then_some(*identifier))
        .collect::<Vec<_>>();
        popup_native::verify_archive_object_locality(
            source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?
                .archive(),
            &archive,
            &changed_identifiers,
        )
        .map_err(|_| Error::Verification)?;
        let expected_archive_bytes = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let archive_bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        if archive_bytes.len() != expected_archive_bytes {
            return Err(Error::Verification);
        }
        budget
            .charge_payload_bytes(archive_bytes.len(), path)
            .map_err(|_| Error::LimitExceeded {
                kind: super::table_cell_pop_up_menu::LimitKind::PayloadBytes,
                observed: archive_bytes.len() as u64,
                maximum: u64::MAX,
                path,
            })?;
        let member_name = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?
            .name()
            .to_owned();
        members.push(native::NativeControlMember {
            member_name,
            archive_bytes,
        });
    }
    Ok(native::NativeControlOutput {
        members,
        component_indices,
        changed_objects: vec![
            (tile_component_index, tile_identifier),
            (format_component_index, format_table_identifier),
        ],
    })
}

fn resolve_custom_cell_graph(
    source: &Package,
    target: CellTarget,
    path: Path,
    for_write: bool,
    budget: &mut TransactionBudget,
) -> Result<CustomCellGraph, Error> {
    let model_component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let model_object =
        native::unique_object(model_component.archive(), target.model_identifier, path)?;
    let model_index = native::unique_message_index(model_object, 6_001, path)?;
    let model_payload = &model_object.messages[model_index].data;
    let (model, model_report) = storage_codec::decode_table_model_with_report(
        model_payload,
        budget.residual_storage_options(model_payload),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, model_report, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model.base_data_store(),
        budget.residual_storage_options(model.base_data_store()),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, store_report, path)?;
    if store.deprecated_custom_format_table().is_some() {
        return Err(Error::UnsupportedDependency { path });
    }
    let tile_options = budget.residual_storage_options(store.tiles());
    let mut tile_visitor = native::TileReferenceCollector::new(budget, path);
    let decoded_tiles = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        tile_options,
        &mut tile_visitor,
    );
    let tile_references = tile_visitor.finish()?;
    let (tile_storage, tile_report) =
        decoded_tiles.map_err(|error| native::map_storage_error(error, path))?;
    native::charge_storage_report(budget, tile_report, path)?;
    let tile_size = tile_storage
        .tile_size()
        .filter(|size| *size != 0)
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target.position.row() / tile_size;
    let mut tile_matches = tile_references
        .iter()
        .filter(|(id, _)| *id == tile_id)
        .map(|(_, reference)| *reference);
    let tile_identifier = tile_matches.next().ok_or(Error::CellNotFound)?;
    if tile_matches.next().is_some()
        || tile_identifier == target.model_identifier
        || tile_references.iter().any(|(_, reference)| *reference == 0)
        || tile_references
            .iter()
            .enumerate()
            .any(|(index, (id, reference))| {
                tile_references[index + 1..]
                    .iter()
                    .any(|(other_id, other_reference)| {
                        id == other_id || reference == other_reference
                    })
            })
    {
        return Err(Error::InvalidSource { path });
    }
    popup::validate_selected_model_reference(source, target, tile_identifier, path)?;
    if for_write {
        native::require_exclusive_selected_model_reference(
            source,
            target,
            tile_identifier,
            path,
            budget,
        )?;
    }
    let tile_component_index = native::resolved_component_index(source, tile_identifier, path)?;
    let tile_archive = source
        .state
        .components
        .catalog()
        .get_index(tile_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let tile_object = native::unique_object(tile_archive, tile_identifier, path)?;
    let tile_message_index = native::unique_message_index(tile_object, 6_002, path)?;
    let tile_payload = native::copy_payload_with_budget(
        &tile_object.messages[tile_message_index].data,
        budget,
        path,
    )?;
    let cell_source = popup_native::tile_cell(
        &tile_payload,
        target.position.row(),
        target.position.column(),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let cell_source = cell_source.to_vec();
    charge_owned_bnc_parse(cell_source.len(), path, budget)?;
    let cell = BncCell::parse(&cell_source).map_err(|_| Error::InvalidSource { path })?;
    native::validate_bnc_format_metadata(&cell_source, &cell, path)?;
    if cell.control_cell_spec_identifier().is_some() {
        return Err(Error::UnsupportedDependency { path });
    }
    let format_identifier = cell.format_identifier();
    if format_identifier.is_some_and(|identifier| identifier == 0) {
        return Err(Error::InvalidSource { path });
    }
    let explicit_flags = cell.explicit_format_flags();
    let kind = cell.cell_format_kind();
    if kind == Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND) {
        let metadata = validate_text_cell_metadata(&cell_source, &cell, path)?;
        if metadata.generic_identifier.is_some() {
            // Converted Text retains a secondary Number edge.  It is not a
            // Custom owner shape; accepting it would require moving two
            // families and could strand the secondary key.
            return Err(Error::UnsupportedDependency { path });
        }
    }
    if format_identifier.is_none() {
        if explicit_flags != 0 {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    if format_table_identifier == target.model_identifier
        || format_table_identifier == tile_identifier
    {
        return Err(Error::UnsupportedDependency { path });
    }
    popup::validate_selected_model_reference(source, target, format_table_identifier, path)?;
    if for_write {
        native::require_exclusive_selected_model_reference(
            source,
            target,
            format_table_identifier,
            path,
            budget,
        )?;
    }
    let format_component_index =
        native::resolved_component_index(source, format_table_identifier, path)?;
    let format_archive = source
        .state
        .components
        .catalog()
        .get_index(format_component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let format_object = native::unique_object(format_archive, format_table_identifier, path)?;
    let format_message =
        native::unique_list_message(format_object, FORMAT_LIST_TYPE, budget, path)?;
    let format_message_index = format_message.message_index;
    let format_payload = native::copy_payload_with_budget(format_message.payload, budget, path)?;
    let format_type = storage_codec::decode_table_data_list_type_with_report(
        &format_payload,
        budget.residual_storage_options(&format_payload),
    )
    .map_err(|error| native::map_storage_error(error, path))?;
    let (list_type, list_report) = format_type;
    native::charge_storage_report(budget, list_report, path)?;
    if list_type.list_type() != FORMAT_LIST_TYPE {
        return Err(Error::UnsupportedDependency { path });
    }
    let facts = native::list_facts(&format_payload, budget, path)?;
    let _ = facts
        .entries
        .iter()
        .find(|entry| Some(entry.key) == format_identifier);
    native::validate_format_refcounts(source, &tile_references, &facts.entries, budget, path)?;
    if let Some(identifier) = format_identifier {
        let entry = facts
            .entries
            .iter()
            .find(|entry| entry.key == identifier)
            .ok_or(Error::InvalidSource { path })?;
        if !entry.is_format || entry.ref_count == 0 || entry.payload.is_empty() {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(CustomCellGraph {
        tile_component_index,
        tile_identifier,
        tile_message_index,
        tile_payload,
        cell_source,
        format_component_index,
        format_table_identifier,
        format_message_index,
        format_payload,
        format_identifier,
        explicit_flags,
        cell_kind: kind,
    })
}

fn selected_custom_reference(
    graph: &CustomCellGraph,
    key: u32,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<CustomReference, Error> {
    let facts = native::list_facts_without_input(&graph.format_payload, budget, path)?;
    let entry = facts
        .entries
        .iter()
        .find(|entry| entry.key == key)
        .ok_or(Error::InvalidSource { path })?;
    if !entry.is_format || entry.ref_count == 0 || entry.payload.is_empty() {
        return Err(Error::InvalidSource { path });
    }
    parse_custom_reference(&entry.payload, budget, path)?
        .ok_or(Error::UnsupportedDependency { path })
}

fn selected_custom_reference_payload(
    graph: &CustomCellGraph,
    key: u32,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let facts = native::list_facts_without_input(&graph.format_payload, budget, path)?;
    let entry = facts
        .entries
        .iter()
        .find(|entry| entry.key == key)
        .ok_or(Error::InvalidSource { path })?;
    if !entry.is_format || entry.ref_count == 0 || entry.payload.is_empty() {
        return Err(Error::InvalidSource { path });
    }
    budget.charge_allocations(1, path)?;
    budget.charge_scratch_bytes(entry.payload.len(), path)?;
    budget.charge_retained_bytes(entry.payload.len(), path)?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(entry.payload.len())
        .map_err(|_| Error::Allocation {
            amount: entry.payload.len(),
            path,
        })?;
    payload.extend_from_slice(&entry.payload);
    Ok(payload)
}

fn custom_reference_matches_cell(format_type: u32, graph: &CustomCellGraph, flags: u16) -> bool {
    matches!(
        (format_type, graph.cell_kind, flags),
        (
            CUSTOM_NUMBER_FORMAT_TYPE,
            Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND),
            litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT,
        ) | (
            CUSTOM_TEXT_FORMAT_TYPE,
            Some(litchi_numbers_wire::TEXT_CELL_FORMAT_KIND),
            litchi_numbers_wire::EXPLICIT_TEXT_FORMAT,
        ) | (
            CUSTOM_DATE_TIME_FORMAT_TYPE,
            Some(litchi_numbers_wire::DATE_TIME_CELL_FORMAT_KIND),
            litchi_numbers_wire::EXPLICIT_DATE_TIME_FORMAT,
        )
    )
}

fn validate_custom_graph_ownership(
    source: &Package,
    target: CellTarget,
    graph: &CustomCellGraph,
    registry: CustomRegistryLocation,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let facts = native::validate_metadata_ownership(
        source,
        &[
            target.component_index,
            graph.tile_component_index,
            graph.format_component_index,
            registry.component_index,
        ],
        budget,
        path,
    )?;
    native::validate_cross_component_reference(
        &facts,
        target.component_index,
        graph.tile_component_index,
        graph.tile_identifier,
        path,
    )?;
    native::validate_cross_component_reference(
        &facts,
        target.component_index,
        graph.format_component_index,
        graph.format_table_identifier,
        path,
    )?;
    let document_component = source
        .state
        .components
        .catalog()
        .iter()
        .position(|component| component.name() == "Index/Document.iwa")
        .ok_or(Error::InvalidSource { path })?;
    native::validate_cross_component_reference(
        &facts,
        document_component,
        registry.component_index,
        registry.object_identifier,
        path,
    )?;
    facts
        .current_uuids_if_registered(&[
            (target.component_index, target.model_identifier),
            (graph.tile_component_index, graph.tile_identifier),
            (graph.format_component_index, graph.format_table_identifier),
            (registry.component_index, registry.object_identifier),
        ])
        .map_err(|_| Error::UnsupportedDependency { path })?;
    Ok(())
}

fn locate_custom_registry(
    source: &Package,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<CustomRegistryLocation, Error> {
    let document_archive = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .ok_or(Error::UnsupportedSource)?;
    let document_object = native::unique_object(document_archive, 1, path)?;
    let document_index =
        native::unique_message_index(document_object, DOCUMENT_MESSAGE_TYPE, path)?;
    let root = &document_object.messages[document_index].data;
    let view = custom_wire_view(root, budget, 0, path)?;
    let mut root_reference = None;
    let mut legacy_reference = None;
    let mut legacy_super_fields = 0usize;
    for field in view.fields() {
        match field.number() {
            CUSTOM_REGISTRY_REFERENCE_FIELD => {
                let reference = parse_custom_registry_reference(field, path)?;
                if root_reference.replace(reference).is_some() {
                    return Err(Error::InvalidSource { path });
                }
            },
            DOCUMENT_LEGACY_SUPER_FIELD => {
                legacy_super_fields = legacy_super_fields
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
                if legacy_super_fields != 1 {
                    return Err(Error::InvalidSource { path });
                }
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource { path });
                }
                field
                    .validate_canonical_framing()
                    .map_err(|_| Error::InvalidSource { path })?;
                let legacy = custom_wire_view(field.payload(), budget, 1, path)?;
                for nested in legacy
                    .fields()
                    .filter(|nested| nested.number() == LEGACY_CUSTOM_REGISTRY_REFERENCE_FIELD)
                {
                    let reference = parse_custom_registry_reference(nested, path)?;
                    if legacy_reference.replace(reference).is_some() {
                        return Err(Error::InvalidSource { path });
                    }
                }
            },
            _ => {},
        }
    }
    let (route, registry_identifier) = match (root_reference, legacy_reference) {
        (Some(registry_identifier), None) => {
            (CustomRegistryRoute::DocumentField9, registry_identifier)
        },
        (None, Some(registry_identifier)) => {
            (CustomRegistryRoute::LegacySuperField12, registry_identifier)
        },
        (Some(_), Some(_)) => return Err(Error::UnsupportedDependency { path }),
        (None, None) => return Err(Error::InvalidSource { path }),
    };
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, registry_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let component_index = resolved.component_index;
    let object = source
        .state
        .components
        .catalog()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.identifier != Some(registry_identifier) || object.messages.len() != 1 {
        return Err(Error::InvalidSource { path });
    }
    let message = object
        .messages
        .first()
        .ok_or(Error::InvalidSource { path })?;
    if message.type_ != CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE {
        return Err(Error::UnsupportedDependency { path });
    }
    for (other_component_index, component) in source.state.components.catalog().iter().enumerate() {
        for other in &component.archive().objects {
            for candidate in &other.messages {
                if other_component_index == component_index
                    && other.archive_info.identifier == Some(registry_identifier)
                    && candidate.type_ == message.type_
                    && candidate.data == message.data
                {
                    continue;
                }
                if candidate.type_ == message.type_
                    && parse_registry_shape(&candidate.data, budget, path).is_ok()
                {
                    return Err(Error::UnsupportedDependency { path });
                }
            }
        }
    }
    Ok(CustomRegistryLocation {
        component_index,
        object_identifier: registry_identifier,
        message_index: 0,
        message_type: message.type_,
        route,
    })
}

fn parse_custom_registry_reference(field: WireFieldView<'_>, path: Path) -> Result<u64, Error> {
    if field.wire_type() != 2 {
        return Err(Error::InvalidSource { path });
    }
    field
        .validate_canonical_framing()
        .map_err(|_| Error::InvalidSource { path })?;
    super::table_headers::resolve::local_reference_identifier(field.payload())
        .map_err(|_| Error::InvalidSource { path })
}

fn registry_payload(
    source: &Package,
    location: CustomRegistryLocation,
    path: Path,
) -> Result<&[u8], Error> {
    let archive = source
        .state
        .components
        .catalog()
        .get_index(location.component_index)
        .ok_or(Error::InvalidSource { path })?
        .archive();
    let object = native::unique_object(archive, location.object_identifier, path)?;
    let message = object
        .messages
        .get(location.message_index)
        .filter(|message| message.type_ == location.message_type)
        .ok_or(Error::InvalidSource { path })?;
    Ok(&message.data)
}

fn parse_registry_shape(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), ()> {
    let fields = custom_wire_view(source, budget, 0, path).map_err(|_| ())?;
    let mut uuids = 0usize;
    let mut formats = 0usize;
    for field in fields.fields() {
        match field.number() {
            CUSTOM_REGISTRY_UUID_FIELD => {
                if field.wire_type() != 2 {
                    return Err(());
                }
                uuids = uuids.checked_add(1).ok_or(())?;
            },
            CUSTOM_REGISTRY_FORMAT_FIELD => {
                if field.wire_type() != 2 {
                    return Err(());
                }
                formats = formats.checked_add(1).ok_or(())?;
            },
            _ => {},
        }
    }
    if uuids == formats { Ok(()) } else { Err(()) }
}

fn parse_custom_registry<'source>(
    source: &'source [u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<CustomRegistryFacts<'source>, Error> {
    let view = custom_wire_view(source, budget, 0, path)?;
    let mut uuid_count = 0usize;
    let mut format_count = 0usize;
    for field in view.fields() {
        match field.number() {
            CUSTOM_REGISTRY_UUID_FIELD | CUSTOM_REGISTRY_FORMAT_FIELD => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource { path });
                }
                field
                    .validate_canonical_framing()
                    .map_err(|_| Error::InvalidSource { path })?;
                let count = if field.number() == CUSTOM_REGISTRY_UUID_FIELD {
                    &mut uuid_count
                } else {
                    &mut format_count
                };
                *count = count.checked_add(1).ok_or(Error::InvalidSource { path })?;
            },
            _ => {},
        }
    }
    if uuid_count != format_count {
        return Err(Error::InvalidSource { path });
    }
    let entry_count = uuid_count;
    let payload_items = entry_count
        .checked_mul(2)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_payload_items(payload_items, path)?;
    budget.charge_payload_references(entry_count, path)?;
    // UUID uniqueness is checked with a bounded linear census below. Charge
    // that worst-case work before entering the loop so a hostile registry
    // cannot turn the no-allocation check into an unbounded CPU path.
    let uniqueness_work = entry_count
        .checked_mul(entry_count.saturating_sub(1))
        .and_then(|work| work.checked_div(2))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_transaction_work(uniqueness_work, path)?;

    let mut uuid_payloads = Vec::new();
    let mut format_payloads = Vec::new();
    let mut entries = Vec::new();
    if entry_count != 0 {
        budget.charge_allocations(3, path)?;
        uuid_payloads
            .try_reserve_exact(entry_count)
            .map_err(|_| Error::Allocation {
                amount: entry_count,
                path,
            })?;
        format_payloads
            .try_reserve_exact(entry_count)
            .map_err(|_| Error::Allocation {
                amount: entry_count,
                path,
            })?;
        entries
            .try_reserve_exact(entry_count)
            .map_err(|_| Error::Allocation {
                amount: entry_count,
                path,
            })?;
    }
    // Keep source payloads borrowed. A rewrite promotes only newly appended
    // records to `Cow::Owned`; existing registry bytes never get cloned just
    // to establish the semantic index.
    for field in view.fields() {
        match field.number() {
            CUSTOM_REGISTRY_UUID_FIELD => uuid_payloads.push(Cow::Borrowed(field.payload())),
            CUSTOM_REGISTRY_FORMAT_FIELD => format_payloads.push(Cow::Borrowed(field.payload())),
            _ => {},
        }
    }
    for (uuid_payload, format_payload) in uuid_payloads.iter().zip(&format_payloads) {
        let uuid = parse_custom_uuid(uuid_payload.as_ref(), budget, path)?;
        if entries
            .iter()
            .any(|entry: &CustomRegistryEntry| entry.uuid == uuid)
        {
            return Err(Error::InvalidSource { path });
        }
        let format = parse_custom_archive(format_payload.as_ref(), budget, path)?;
        if custom_format_type(&format) == 0 {
            return Err(Error::InvalidSource { path });
        }
        entries.push(CustomRegistryEntry {
            uuid,
            format: Some(format),
        });
    }
    Ok(CustomRegistryFacts {
        source,
        uuid_payloads,
        format_payloads,
        entries,
    })
}

fn rewrite_custom_registry(
    registry: &CustomRegistryFacts<'_>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let mut output = wire_rewrite_repeated(
        registry.source,
        CUSTOM_REGISTRY_UUID_FIELD,
        &registry.uuid_payloads,
        path,
        budget,
    )?;
    output = wire_rewrite_repeated(
        &output,
        CUSTOM_REGISTRY_FORMAT_FIELD,
        &registry.format_payloads,
        path,
        budget,
    )?;
    Ok(output)
}

fn wire_rewrite_repeated<'source>(
    source: &[u8],
    field_number: u32,
    replacements: &[Cow<'source, [u8]>],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let view = custom_wire_view(source, budget, 0, path)?;
    let mut matched_count = 0usize;
    let mut removed_bytes = 0usize;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 || field.validate_canonical_framing().is_err() {
            return Err(Error::InvalidSource { path });
        }
        matched_count = matched_count
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
        removed_bytes = removed_bytes
            .checked_add(field.raw().len())
            .ok_or(Error::InvalidSource { path })?;
    }
    let retained_bytes = source
        .len()
        .checked_sub(removed_bytes)
        .ok_or(Error::InvalidSource { path })?;
    let replacement_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        let replacement = replacement.as_ref();
        let key = (u64::from(field_number) << 3) | 2;
        total
            .checked_add(litchi_iwa_common::varint::encoded_len(key))
            .and_then(|total| {
                total.checked_add(litchi_iwa_common::varint::encoded_len(
                    replacement.len() as u64
                ))
            })
            .and_then(|total| total.checked_add(replacement.len()))
            .ok_or(Error::InvalidSource { path })
    })?;
    let output_len = retained_bytes
        .checked_add(replacement_bytes)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_output(output_len, path)?;
    budget.charge_allocations(1, path)?;
    budget.charge_scratch_bytes(output_len, path)?;
    budget.charge_retained_bytes(output_len, path)?;
    budget.charge_transaction_work(output_len, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    let mut matched = 0usize;
    for field in view.fields() {
        if field.number() != field_number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if let Some(replacement) = replacements.get(matched) {
            let replacement = replacement.as_ref();
            output.extend_from_slice(field.key());
            litchi_iwa_common::varint::encode_varint_into(
                &mut output,
                u64::try_from(replacement.len()).map_err(|_| Error::InvalidSource { path })?,
            );
            output.extend_from_slice(replacement);
        }
        matched = matched.saturating_add(1);
    }
    for replacement in replacements.iter().skip(matched_count) {
        let key = (u64::from(field_number) << 3) | 2;
        let replacement = replacement.as_ref();
        litchi_iwa_common::varint::encode_varint_into(&mut output, key);
        litchi_iwa_common::varint::encode_varint_into(
            &mut output,
            u64::try_from(replacement.len()).map_err(|_| Error::InvalidSource { path })?,
        );
        output.extend_from_slice(replacement);
    }
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn custom_wire_view<'source>(
    source: &'source [u8],
    budget: &mut TransactionBudget,
    depth: u32,
    path: Path,
) -> Result<WireView<'source>, Error> {
    budget.charge_wire_bytes(source.len(), path)?;
    budget.charge_wire_work(source.len(), path)?;
    budget.charge_wire_nesting(depth, path)?;
    let view = WireView::parse(source).map_err(|_| Error::InvalidSource { path })?;
    budget.charge_wire_fields(view.len(), path)?;
    Ok(view)
}

fn parse_custom_uuid(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<CustomUuid, Error> {
    let view = custom_wire_view(source, budget, 1, path)?;
    let lower = parse_required_varint(&view, 1, path)?;
    let upper = parse_required_varint(&view, 2, path)?;
    if lower == 0 || upper == 0 {
        return Err(Error::InvalidSource { path });
    }
    budget.charge_wire_reference_bytes(source.len(), path)?;
    Ok(CustomUuid { lower, upper })
}

fn parse_required_varint(view: &WireView<'_>, number: u32, path: Path) -> Result<u64, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 0 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_key()
            .map_err(|_| Error::InvalidSource { path })?;
        let payload = field.payload();
        let (value, width) =
            decode_varint_from_bytes(payload).map_err(|_| Error::InvalidSource { path })?;
        if width != payload.len() || litchi_iwa_common::varint::encoded_len(value) != width {
            return Err(Error::InvalidSource { path });
        }
        found = Some(value);
    }
    found.ok_or(Error::InvalidSource { path })
}

fn parse_optional_varint(
    view: &WireView<'_>,
    number: u32,
    path: Path,
) -> Result<Option<u64>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 0 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_key()
            .map_err(|_| Error::InvalidSource { path })?;
        let payload = field.payload();
        let (value, width) =
            decode_varint_from_bytes(payload).map_err(|_| Error::InvalidSource { path })?;
        if width != payload.len() || litchi_iwa_common::varint::encoded_len(value) != width {
            return Err(Error::InvalidSource { path });
        }
        found = Some(value);
    }
    Ok(found)
}

fn parse_optional_bool(
    view: &WireView<'_>,
    number: u32,
    path: Path,
) -> Result<Option<bool>, Error> {
    parse_optional_varint(view, number, path)?.map_or(Ok(None), |value| match value {
        0 => Ok(Some(false)),
        1 => Ok(Some(true)),
        _ => Err(Error::InvalidSource { path }),
    })
}

fn parse_required_string(
    view: &WireView<'_>,
    number: u32,
    maximum: usize,
    reject_surrounding_whitespace: bool,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<String, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource { path })?;
        let payload = field.payload();
        let value = std::str::from_utf8(payload).map_err(|_| Error::InvalidSource { path })?;
        // Validate the borrowed bytes and debit their text budget before a
        // `String` is materialized. This keeps hostile oversized or control
        // laden names/patterns fail-closed without an attacker-sized heap
        // allocation, even though semantic constructors validate again.
        if value.is_empty()
            || value.len() > maximum
            || value.chars().any(char::is_control)
            || (reject_surrounding_whitespace && value.trim() != value)
        {
            return Err(Error::InvalidSource { path });
        }
        budget.charge_wire_text_bytes(value.len(), path)?;
        budget.charge_allocations(1, path)?;
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| Error::Allocation {
                amount: value.len(),
                path,
            })?;
        owned.push_str(value);
        let value = owned;
        found = Some(value);
    }
    found.ok_or(Error::InvalidSource { path })
}

fn parse_required_bytes<'source>(
    view: &WireView<'source>,
    number: u32,
    path: Path,
) -> Result<&'source [u8], Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource { path })?;
        found = Some(field.payload());
    }
    found.ok_or(Error::InvalidSource { path })
}

fn parse_fixed32(view: &WireView<'_>, number: u32, path: Path) -> Result<Option<u32>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 5 || field.payload().len() != 4 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_key()
            .map_err(|_| Error::InvalidSource { path })?;
        found = Some(u32::from_le_bytes(
            field
                .payload()
                .try_into()
                .map_err(|_| Error::InvalidSource { path })?,
        ));
    }
    Ok(found)
}

fn parse_fixed64(view: &WireView<'_>, number: u32, path: Path) -> Result<Option<u64>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.is_some() || field.wire_type() != 1 || field.payload().len() != 8 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_key()
            .map_err(|_| Error::InvalidSource { path })?;
        found = Some(u64::from_le_bytes(
            field
                .payload()
                .try_into()
                .map_err(|_| Error::InvalidSource { path })?,
        ));
    }
    Ok(found)
}

fn custom_format_type(value: &Custom) -> u32 {
    match value {
        Custom::Number(_) => CUSTOM_NUMBER_FORMAT_TYPE,
        Custom::Text(_) => CUSTOM_TEXT_FORMAT_TYPE,
        Custom::DateTime(_) => CUSTOM_DATE_TIME_FORMAT_TYPE,
    }
}

fn parse_custom_archive(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Custom, Error> {
    let view = custom_wire_view(source, budget, 1, path)?;
    let name = Name::from_owned(parse_required_string(
        &view,
        CUSTOM_FORMAT_NAME_FIELD,
        MAX_NAME_BYTES,
        true,
        budget,
        path,
    )?)
    .map_err(|_| Error::InvalidSource { path })?;
    let pre_type = u32::try_from(parse_required_varint(
        &view,
        CUSTOM_FORMAT_PRE_TYPE_FIELD,
        path,
    )?)
    .map_err(|_| Error::InvalidSource { path })?;
    if !matches!(
        pre_type,
        CUSTOM_NUMBER_FORMAT_TYPE | CUSTOM_TEXT_FORMAT_TYPE | CUSTOM_DATE_TIME_FORMAT_TYPE
    ) {
        return Err(Error::UnsupportedDependency { path });
    }
    let declared_type = parse_optional_varint(&view, CUSTOM_FORMAT_TYPE_FIELD, path)?;
    if declared_type.is_some_and(|value| value != u64::from(pre_type)) {
        return Err(Error::InvalidSource { path });
    }
    let default_payload = parse_required_bytes(&view, CUSTOM_FORMAT_DEFAULT_FIELD, path)?;
    let (_, default_pattern) = parse_custom_pattern(default_payload, pre_type, budget, path)?;
    let condition_count = view
        .fields()
        .filter(|field| field.number() == CUSTOM_FORMAT_CONDITION_FIELD)
        .count();
    if condition_count > crate::cell::data_format::custom::MAX_RULES {
        return Err(Error::UnsupportedDependency { path });
    }
    let mut conditions = Vec::new();
    if condition_count != 0 {
        budget.charge_allocations(1, path)?;
        conditions
            .try_reserve_exact(condition_count)
            .map_err(|_| Error::Allocation {
                amount: condition_count,
                path,
            })?;
    }
    for field in view
        .fields()
        .filter(|field| field.number() == CUSTOM_FORMAT_CONDITION_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource { path })?;
        conditions.push(parse_custom_condition(
            field.payload(),
            pre_type,
            budget,
            path,
        )?);
    }
    match pre_type {
        CUSTOM_NUMBER_FORMAT_TYPE => {
            let pattern = NumberPattern::from_owned(default_pattern)
                .map_err(|_| Error::InvalidSource { path })?;
            let mut rules = Vec::new();
            if !conditions.is_empty() {
                budget.charge_allocations(1, path)?;
                rules
                    .try_reserve_exact(conditions.len())
                    .map_err(|_| Error::Allocation {
                        amount: conditions.len(),
                        path,
                    })?;
            }
            for (condition, pattern) in conditions {
                let pattern = NumberPattern::from_owned(pattern)
                    .map_err(|_| Error::InvalidSource { path })?;
                rules.push(NumberRule::new(condition, pattern));
            }
            Ok(Custom::Number(
                CustomNumber::try_with_rules(name, pattern, rules)
                    .map_err(|_| Error::InvalidSource { path })?,
            ))
        },
        CUSTOM_TEXT_FORMAT_TYPE => {
            if !conditions.is_empty() {
                return Err(Error::UnsupportedDependency { path });
            }
            let (prefix, suffix, includes_cell) =
                split_custom_text_pattern(default_pattern, budget, path)?;
            let value = if includes_cell {
                CustomText::try_new(name, prefix, suffix)
            } else {
                CustomText::try_literal(name, prefix)
            }
            .map_err(|_| Error::InvalidSource { path })?;
            Ok(Custom::Text(value))
        },
        CUSTOM_DATE_TIME_FORMAT_TYPE => {
            if !conditions.is_empty() {
                return Err(Error::UnsupportedDependency { path });
            }
            let pattern = DateTimePattern::from_owned(default_pattern)
                .map_err(|_| Error::InvalidSource { path })?;
            Ok(Custom::DateTime(CustomDateTime::new(name, pattern)))
        },
        _ => Err(Error::UnsupportedDependency { path }),
    }
}

fn parse_custom_condition(
    source: &[u8],
    expected_type: u32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(Condition, String), Error> {
    let view = custom_wire_view(source, budget, 2, path)?;
    let condition_type = u32::try_from(parse_required_varint(
        &view,
        CUSTOM_CONDITION_TYPE_FIELD,
        path,
    )?)
    .map_err(|_| Error::InvalidSource { path })?;
    let fixed = parse_fixed32(&view, CUSTOM_CONDITION_FLOAT_FIELD, path)?;
    let double = parse_fixed64(&view, CUSTOM_CONDITION_DOUBLE_FIELD, path)?;
    if fixed.is_some() == double.is_some() {
        return Err(Error::InvalidSource { path });
    }
    let threshold = if let Some(value) = double {
        f64::from_bits(value)
    } else {
        f32::from_bits(fixed.ok_or(Error::InvalidSource { path })?) as f64
    };
    let threshold =
        ConditionValue::try_new(threshold).map_err(|_| Error::InvalidSource { path })?;
    let condition = match condition_type {
        0 => Condition::EqualTo(threshold),
        1 => Condition::LessThan(threshold),
        2 => Condition::LessThanOrEqualTo(threshold),
        3 => Condition::GreaterThan(threshold),
        4 => Condition::GreaterThanOrEqualTo(threshold),
        _ => return Err(Error::UnsupportedDependency { path }),
    };
    let format_payload = parse_required_bytes(&view, CUSTOM_CONDITION_FORMAT_FIELD, path)?;
    let (format_type, pattern) = parse_custom_pattern(format_payload, expected_type, budget, path)?;
    if format_type != expected_type {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok((condition, pattern))
}

fn parse_custom_pattern(
    source: &[u8],
    expected_type: u32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(u32, String), Error> {
    let view = custom_wire_view(source, budget, 2, path)?;
    let format_type = u32::try_from(parse_required_varint(
        &view,
        CUSTOM_PATTERN_TYPE_FIELD,
        path,
    )?)
    .map_err(|_| Error::InvalidSource { path })?;
    if format_type != expected_type {
        return Err(Error::UnsupportedDependency { path });
    }
    let pattern = parse_required_string(
        &view,
        CUSTOM_PATTERN_STRING_FIELD,
        MAX_PATTERN_BYTES,
        false,
        budget,
        path,
    )?;
    // The custom pattern is a strict FormatStructArchive shape. Optional
    // fields are accepted only when they carry one of the known native or
    // source-built custom-format cache profiles; all other known fields
    // belong to a different owner and are rejected.
    let show_thousands = parse_optional_bool(&view, 5, path)?;
    if expected_type != CUSTOM_NUMBER_FORMAT_TYPE && show_thousands.is_some_and(|value| value) {
        return Err(Error::UnsupportedDependency { path });
    }
    if parse_optional_bool(&view, 6, path)?.is_some_and(|value| value)
        || parse_optional_varint(&view, 11, path)?
            .is_some_and(|value| value != u64::from(CUSTOM_FRACTION_SENTINEL))
        || parse_optional_fixed64_one(&view, 19, path)? == Some(false)
        || parse_optional_bool(&view, 20, path)?.is_some_and(|value| value)
        || parse_optional_bool(&view, 36, path)?.is_some_and(|value| value)
    {
        return Err(Error::UnsupportedDependency { path });
    }
    for field_number in [27_u32, 28, 29, 30, 34, 35] {
        if parse_optional_varint(&view, field_number, path)?.is_some_and(|value| value != 0) {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    // Numbers uses field 31 as a cached pattern index. Source-built Number
    // formats store the UTF-16 width after their final integer placeholder;
    // native-authored Text stores the final UTF-16 code-unit index, while the
    // source-built Text writer stores the UTF-16 suffix width after its unique
    // value token. The zero value remains the canonical empty-cache profile.
    // Admit only these derived values and keep every other display family on
    // its previous zero-only profile.
    let index_from_right_last_integer = parse_optional_varint(&view, 31, path)?;
    if let Some(value) = index_from_right_last_integer.filter(|value| *value != 0) {
        let expected_index = match expected_type {
            CUSTOM_NUMBER_FORMAT_TYPE => custom_number_pattern_suffix_width(&pattern),
            CUSTOM_TEXT_FORMAT_TYPE => pattern
                .encode_utf16()
                .count()
                .checked_sub(1)
                .and_then(|index| u64::try_from(index).ok())
                .ok_or(Error::InvalidSource { path }),
            _ => Err(Error::UnsupportedDependency { path }),
        }?;
        if value != expected_index
            && (expected_type != CUSTOM_TEXT_FORMAT_TYPE
                || Some(value) != custom_text_pattern_suffix_width(&pattern))
        {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    if let Some(contains_integer) = parse_optional_bool(&view, 37, path)? {
        // The focused source-built writer stores false, while the legacy host
        // custom-format writer stores true for Date & Time. The bool parser
        // already bounds this compatibility profile to the two canonical
        // wire values.
        let valid = if expected_type == CUSTOM_DATE_TIME_FORMAT_TYPE {
            true
        } else {
            let expected = expected_type == CUSTOM_NUMBER_FORMAT_TYPE
                && pattern
                    .chars()
                    .any(|character| matches!(character, '#' | '0'));
            contains_integer == expected
        };
        if !valid {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    for field in view.fields() {
        if field.number() > 45
            || matches!(
                field.number(),
                1 | 5 | 6 | 11 | 18 | 19 | 20 | 27 | 28 | 29 | 30 | 31 | 34 | 35 | 36 | 37
            )
        {
            continue;
        }
        return Err(Error::UnsupportedDependency { path });
    }
    Ok((format_type, pattern))
}

fn parse_optional_fixed64_one(
    view: &WireView<'_>,
    number: u32,
    path: Path,
) -> Result<Option<bool>, Error> {
    Ok(parse_fixed64(view, number, path)?.map(|bits| f64::from_bits(bits) == 1.0))
}

fn split_custom_text_pattern(
    pattern: String,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(String, String, bool), Error> {
    let Some((prefix, suffix)) = pattern.split_once(CUSTOM_TEXT_VALUE_TOKEN) else {
        if pattern.is_empty() {
            return Err(Error::InvalidSource { path });
        }
        // The owned pattern can be adopted directly by the literal Text
        // constructor. This avoids a second pattern-sized clone.
        return Ok((pattern, String::new(), false));
    };
    if suffix.contains(CUSTOM_TEXT_VALUE_TOKEN) {
        return Err(Error::UnsupportedDependency { path });
    }
    budget.charge_allocations(2, path)?;
    let mut prefix_owned = String::new();
    prefix_owned
        .try_reserve_exact(prefix.len())
        .map_err(|_| Error::Allocation {
            amount: prefix.len(),
            path,
        })?;
    prefix_owned.push_str(prefix);
    let mut suffix_owned = String::new();
    suffix_owned
        .try_reserve_exact(suffix.len())
        .map_err(|_| Error::Allocation {
            amount: suffix.len(),
            path,
        })?;
    suffix_owned.push_str(suffix);
    Ok((prefix_owned, suffix_owned, true))
}

fn custom_varint_field_len(field: u32, value: u64) -> usize {
    litchi_iwa_common::varint::encoded_len(u64::from(field) << 3)
        .saturating_add(litchi_iwa_common::varint::encoded_len(value))
}

fn custom_bytes_field_len(field: u32, payload_len: usize) -> Result<usize, Error> {
    let payload_len = u64::try_from(payload_len).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })?;
    custom_varint_field_len(field, payload_len)
        .checked_add(payload_len as usize)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

fn custom_fixed64_field_len(field: u32) -> usize {
    litchi_iwa_common::varint::encoded_len((u64::from(field) << 3) | 1) + 8
}

fn custom_pattern_encoded_len(format_type: u32, pattern: &str) -> Result<usize, Error> {
    let index_from_right_last_integer =
        custom_pattern_index_from_right_last_integer(format_type, pattern)?;
    let mut length = 0usize;
    // All constant pattern fields are emitted canonically as varints. Their
    // values are included here (rather than using a blanket overhead) so the
    // reservation remains an exact preflight for the fallible output Vec.
    for (field, value) in [
        (CUSTOM_PATTERN_TYPE_FIELD, u64::from(format_type)),
        (
            5,
            u64::from(format_type == CUSTOM_NUMBER_FORMAT_TYPE && pattern.contains(',')),
        ),
        (6, 0),
        (11, u64::from(CUSTOM_FRACTION_SENTINEL)),
        (20, 0),
        (27, 0),
        (28, 0),
        (29, 0),
        (30, 0),
        (31, index_from_right_last_integer),
        (34, 0),
        (35, 0),
        (36, 0),
        (
            37,
            u64::from(
                format_type == CUSTOM_NUMBER_FORMAT_TYPE
                    && pattern
                        .chars()
                        .any(|character| matches!(character, '#' | '0')),
            ),
        ),
    ] {
        length = length
            .checked_add(custom_varint_field_len(field, value))
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    }
    let pattern_field_len = custom_bytes_field_len(CUSTOM_PATTERN_STRING_FIELD, pattern.len())?;
    length = length
        .checked_add(pattern_field_len)
        .and_then(|length| length.checked_add(custom_fixed64_field_len(19)))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    Ok(length)
}

fn custom_pattern_index_from_right_last_integer(
    format_type: u32,
    pattern: &str,
) -> Result<u64, Error> {
    match format_type {
        CUSTOM_NUMBER_FORMAT_TYPE => custom_number_pattern_suffix_width(pattern),
        CUSTOM_TEXT_FORMAT_TYPE => pattern
            .encode_utf16()
            .count()
            .checked_sub(1)
            .and_then(|index| u64::try_from(index).ok())
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            }),
        _ => Ok(0),
    }
}

fn custom_number_pattern_suffix_width(pattern: &str) -> Result<u64, Error> {
    let suffix = pattern
        .char_indices()
        .rev()
        .find(|(_, character)| matches!(character, '#' | '0'))
        .map_or(pattern, |(offset, character)| {
            &pattern[offset + character.len_utf8()..]
        });
    u64::try_from(suffix.encode_utf16().count()).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })
}

fn custom_text_pattern_suffix_width(pattern: &str) -> Option<u64> {
    let (_, suffix) = pattern.split_once(CUSTOM_TEXT_VALUE_TOKEN)?;
    if suffix.contains(CUSTOM_TEXT_VALUE_TOKEN) {
        return None;
    }
    u64::try_from(suffix.encode_utf16().count()).ok()
}

fn custom_condition_encoded_len(format_type: u32, pattern: &str) -> Result<usize, Error> {
    let pattern_len = custom_pattern_encoded_len(format_type, pattern)?;
    let condition_format_len = custom_bytes_field_len(CUSTOM_CONDITION_FORMAT_FIELD, pattern_len)?;
    custom_varint_field_len(CUSTOM_CONDITION_TYPE_FIELD, 0)
        .checked_add(custom_fixed64_field_len(CUSTOM_CONDITION_DOUBLE_FIELD))
        .and_then(|length| length.checked_add(condition_format_len))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

fn custom_archive_encoded_len(
    name_len: usize,
    format_type: u32,
    default_pattern: &str,
    rules: Option<&[NumberRule]>,
) -> Result<usize, Error> {
    let default_pattern_len = custom_pattern_encoded_len(format_type, default_pattern)?;
    let name_field_len = custom_bytes_field_len(CUSTOM_FORMAT_NAME_FIELD, name_len)?;
    let default_field_len =
        custom_bytes_field_len(CUSTOM_FORMAT_DEFAULT_FIELD, default_pattern_len)?;
    let mut length = name_field_len
        .checked_add(custom_varint_field_len(
            CUSTOM_FORMAT_PRE_TYPE_FIELD,
            u64::from(format_type),
        ))
        .and_then(|length| length.checked_add(default_field_len))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if let Some(rules) = rules {
        for rule in rules {
            let condition_len = custom_condition_encoded_len(format_type, rule.pattern().as_str())?;
            length = length
                .checked_add(custom_bytes_field_len(
                    CUSTOM_FORMAT_CONDITION_FIELD,
                    condition_len,
                )?)
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
        }
    }
    length = length
        .checked_add(custom_varint_field_len(
            CUSTOM_FORMAT_TYPE_FIELD,
            u64::from(format_type),
        ))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    Ok(length)
}

fn encode_custom_registry_entry(
    uuid: CustomUuid,
    value: &Custom,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(Vec<u8>, Vec<u8>), Error> {
    let uuid_payload = encode_custom_uuid(uuid, path, budget)?;
    let format_payload = encode_custom_archive(value, path, budget)?;
    Ok((uuid_payload, format_payload))
}

fn encode_custom_archive(
    value: &Custom,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let (name, format_type, default_pattern, rules): (
        &str,
        u32,
        Cow<'_, str>,
        Option<&[NumberRule]>,
    ) = match value {
        Custom::Number(value) => (
            value.name().as_str(),
            CUSTOM_NUMBER_FORMAT_TYPE,
            Cow::Borrowed(value.default_pattern().as_str()),
            Some(value.rules()),
        ),
        Custom::Text(value) => {
            let default_pattern = if value.includes_cell_text() {
                let length = value
                    .prefix()
                    .len()
                    .checked_add(value.suffix().len())
                    .and_then(|length| length.checked_add(CUSTOM_TEXT_VALUE_TOKEN.len_utf8()))
                    .ok_or(Error::InvalidSource { path })?;
                budget.charge_allocations(1, path)?;
                let mut pattern = String::new();
                pattern
                    .try_reserve_exact(length)
                    .map_err(|_| Error::Allocation {
                        amount: length,
                        path,
                    })?;
                pattern.push_str(value.prefix());
                pattern.push(CUSTOM_TEXT_VALUE_TOKEN);
                pattern.push_str(value.suffix());
                Cow::Owned(pattern)
            } else {
                Cow::Borrowed(value.prefix())
            };
            (
                value.name().as_str(),
                CUSTOM_TEXT_FORMAT_TYPE,
                default_pattern,
                None,
            )
        },
        Custom::DateTime(value) => (
            value.name().as_str(),
            CUSTOM_DATE_TIME_FORMAT_TYPE,
            Cow::Borrowed(value.pattern().as_str()),
            None,
        ),
    };
    let output_len =
        custom_archive_encoded_len(name.len(), format_type, default_pattern.as_ref(), rules)?;
    budget.charge_allocations(1, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    append_custom_string(&mut output, CUSTOM_FORMAT_NAME_FIELD, name, path)?;
    append_custom_varint(
        &mut output,
        CUSTOM_FORMAT_PRE_TYPE_FIELD,
        u64::from(format_type),
        path,
    )?;
    let pattern = encode_custom_pattern(format_type, default_pattern.as_ref(), path, budget)?;
    append_custom_bytes(&mut output, CUSTOM_FORMAT_DEFAULT_FIELD, &pattern, path)?;
    if let Some(rules) = rules {
        for rule in rules {
            let condition = encode_custom_condition(rule, format_type, path, budget)?;
            append_custom_bytes(&mut output, CUSTOM_FORMAT_CONDITION_FIELD, &condition, path)?;
        }
    }
    append_custom_varint(
        &mut output,
        CUSTOM_FORMAT_TYPE_FIELD,
        u64::from(format_type),
        path,
    )?;
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn encode_custom_pattern(
    format_type: u32,
    pattern: &str,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_len = custom_pattern_encoded_len(format_type, pattern)?;
    budget.charge_allocations(1, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    append_custom_varint(
        &mut output,
        CUSTOM_PATTERN_TYPE_FIELD,
        u64::from(format_type),
        path,
    )?;
    append_custom_varint(
        &mut output,
        5,
        u64::from(format_type == CUSTOM_NUMBER_FORMAT_TYPE && pattern.contains(',')),
        path,
    )?;
    append_custom_varint(&mut output, 6, 0, path)?;
    append_custom_varint(&mut output, 11, u64::from(CUSTOM_FRACTION_SENTINEL), path)?;
    append_custom_string(&mut output, CUSTOM_PATTERN_STRING_FIELD, pattern, path)?;
    append_custom_fixed64(&mut output, 19, 1.0_f64.to_bits(), path)?;
    append_custom_varint(&mut output, 20, 0, path)?;
    for field in [27_u32, 28, 29, 30] {
        append_custom_varint(&mut output, field, 0, path)?;
    }
    append_custom_varint(
        &mut output,
        31,
        custom_pattern_index_from_right_last_integer(format_type, pattern)?,
        path,
    )?;
    for field in [34_u32, 35] {
        append_custom_varint(&mut output, field, 0, path)?;
    }
    append_custom_varint(&mut output, 36, 0, path)?;
    append_custom_varint(
        &mut output,
        37,
        u64::from(
            format_type == CUSTOM_NUMBER_FORMAT_TYPE
                && pattern
                    .chars()
                    .any(|character| matches!(character, '#' | '0')),
        ),
        path,
    )?;
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn encode_custom_condition(
    rule: &NumberRule,
    format_type: u32,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_len = custom_condition_encoded_len(format_type, rule.pattern().as_str())?;
    budget.charge_allocations(1, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    let condition_type = match rule.condition() {
        Condition::EqualTo(_) => 0,
        Condition::LessThan(_) => 1,
        Condition::LessThanOrEqualTo(_) => 2,
        Condition::GreaterThan(_) => 3,
        Condition::GreaterThanOrEqualTo(_) => 4,
    };
    append_custom_varint(
        &mut output,
        CUSTOM_CONDITION_TYPE_FIELD,
        condition_type,
        path,
    )?;
    append_custom_fixed64(
        &mut output,
        CUSTOM_CONDITION_DOUBLE_FIELD,
        rule.condition().threshold().value().to_bits(),
        path,
    )?;
    let pattern = encode_custom_pattern(format_type, rule.pattern().as_str(), path, budget)?;
    append_custom_bytes(&mut output, CUSTOM_CONDITION_FORMAT_FIELD, &pattern, path)?;
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn append_custom_varint(
    output: &mut Vec<u8>,
    field: u32,
    value: u64,
    path: Path,
) -> Result<(), Error> {
    if field == 0 || field > 0x1fff_ffff {
        return Err(Error::InvalidSource { path });
    }
    litchi_iwa_common::varint::encode_varint_into(output, u64::from(field) << 3);
    litchi_iwa_common::varint::encode_varint_into(output, value);
    Ok(())
}

fn append_custom_bytes(
    output: &mut Vec<u8>,
    field: u32,
    value: &[u8],
    path: Path,
) -> Result<(), Error> {
    if field == 0 || field > 0x1fff_ffff {
        return Err(Error::InvalidSource { path });
    }
    let length = u64::try_from(value.len()).map_err(|_| Error::InvalidSource { path })?;
    litchi_iwa_common::varint::encode_varint_into(output, (u64::from(field) << 3) | 2);
    litchi_iwa_common::varint::encode_varint_into(output, length);
    output.extend_from_slice(value);
    Ok(())
}

fn append_custom_string(
    output: &mut Vec<u8>,
    field: u32,
    value: &str,
    path: Path,
) -> Result<(), Error> {
    append_custom_bytes(output, field, value.as_bytes(), path)
}

fn append_custom_fixed64(
    output: &mut Vec<u8>,
    field: u32,
    value: u64,
    _path: Path,
) -> Result<(), Error> {
    if field == 0 || field > 0x1fff_ffff {
        return Err(Error::InvalidSource { path: _path });
    }
    let key = (u64::from(field) << 3) | 1;
    litchi_iwa_common::varint::encode_varint_into(output, key);
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn encode_custom_uuid(
    uuid: CustomUuid,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let output_len = custom_varint_field_len(1, uuid.lower)
        .checked_add(custom_varint_field_len(2, uuid.upper))
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(1, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    append_custom_varint(&mut output, 1, uuid.lower, path)?;
    append_custom_varint(&mut output, 2, uuid.upper, path)?;
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn encode_custom_reference(
    format_type: u32,
    uuid: CustomUuid,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let uuid_len = custom_varint_field_len(1, uuid.lower)
        .checked_add(custom_varint_field_len(2, uuid.upper))
        .ok_or(Error::InvalidSource { path })?;
    let output_len = custom_varint_field_len(CUSTOM_REFERENCE_TYPE_FIELD, u64::from(format_type))
        .checked_add(custom_bytes_field_len(
            CUSTOM_REFERENCE_UUID_FIELD,
            uuid_len,
        )?)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_allocations(1, path)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation {
            amount: output_len,
            path,
        })?;
    append_custom_varint(
        &mut output,
        CUSTOM_REFERENCE_TYPE_FIELD,
        u64::from(format_type),
        path,
    )?;
    let uuid_payload = encode_custom_uuid(uuid, path, budget)?;
    append_custom_bytes(
        &mut output,
        CUSTOM_REFERENCE_UUID_FIELD,
        &uuid_payload,
        path,
    )?;
    if output.len() != output_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(output)
}

fn rewrite_custom_reference(
    source: &[u8],
    format_type: u32,
    uuid: CustomUuid,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    // Both common wire patch helpers allocate a complete replacement buffer.
    // Debit the first reservation before invoking it; the UUID payload has
    // its own charge in `encode_custom_uuid`, and the second patch is charged
    // immediately before its allocation below.
    budget.charge_allocations(1, path)?;
    let output = litchi_iwa_common::wire::patch_varint_field(
        source,
        CUSTOM_REFERENCE_TYPE_FIELD,
        true,
        Some(u64::from(format_type)),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let uuid_payload = encode_custom_uuid(uuid, path, budget)?;
    budget.charge_allocations(1, path)?;
    let output = litchi_iwa_common::wire::patch_length_delimited_field(
        &output,
        CUSTOM_REFERENCE_UUID_FIELD,
        true,
        Some(&uuid_payload),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    Ok(output)
}

fn rewrite_custom_cell_clear(
    source: &[u8],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    charge_owned_bnc_parse(source.len(), path, budget)?;
    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    cell.clear_explicit_format();
    cell.try_encode_with_limit(
        source
            .len()
            .checked_add(16)
            .ok_or(Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })
}

fn fresh_custom_uuid(entries: &[CustomRegistryEntry]) -> CustomUuid {
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let uuid = CustomUuid {
            lower: u64::from_le_bytes(bytes[..8].try_into().expect("GUID lower width")),
            upper: u64::from_le_bytes(bytes[8..].try_into().expect("GUID upper width")),
        };
        if uuid.lower != 0 && uuid.upper != 0 && !entries.iter().any(|entry| entry.uuid == uuid) {
            return uuid;
        }
    }
}

fn custom_uuid_is_referenced(
    source: &Package,
    replacement_component: usize,
    replacement_identifier: u64,
    replacement_message: usize,
    replacement_payload: &[u8],
    needle: CustomUuid,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<bool, Error> {
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for object in &component.archive().objects {
            let object_identifier = object.archive_info.identifier.unwrap_or(0);
            for (message_index, message) in object.messages.iter().enumerate() {
                let payload = if component_index == replacement_component
                    && object_identifier == replacement_identifier
                    && message_index == replacement_message
                {
                    replacement_payload
                } else {
                    &message.data
                };
                if message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE {
                    let (list_type, report) =
                        storage_codec::decode_table_data_list_type_with_report(
                            payload,
                            budget.residual_storage_options(payload),
                        )
                        .map_err(|error| native::map_storage_error(error, path))?;
                    native::charge_storage_report(budget, report, path)?;
                    if list_type.list_type() == CUSTOM_FORMAT_LIST_TYPE {
                        return Err(Error::UnsupportedDependency { path });
                    }
                    if list_type.list_type() != FORMAT_LIST_TYPE {
                        continue;
                    }
                    let facts = if component_index == replacement_component
                        && object_identifier == replacement_identifier
                        && message_index == replacement_message
                    {
                        native::list_facts_without_input(payload, budget, path)?
                    } else {
                        native::list_facts(payload, budget, path)?
                    };
                    for entry in facts.entries.iter().filter(|entry| entry.is_format) {
                        if let Some(reference) =
                            parse_custom_reference(&entry.payload, budget, path)?
                        {
                            if reference.uuid == needle {
                                return Ok(true);
                            }
                        }
                    }
                } else if message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE {
                    let (list_type, report) =
                        storage_codec::decode_table_data_list_segment_type_with_report(
                            payload,
                            budget.residual_storage_options(payload),
                        )
                        .map_err(|error| native::map_storage_error(error, path))?;
                    native::charge_storage_report(budget, report, path)?;
                    if matches!(
                        list_type.list_type(),
                        CUSTOM_FORMAT_LIST_TYPE | FORMAT_LIST_TYPE
                    ) {
                        return Err(Error::UnsupportedDependency { path });
                    }
                }
            }
        }
    }
    Ok(false)
}

fn parse_custom_reference(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Option<CustomReference>, Error> {
    let view = custom_wire_view(source, budget, 1, path)?;
    let format_type = parse_optional_varint(&view, CUSTOM_REFERENCE_TYPE_FIELD, path)?;
    let mut uuid_payload = None;
    let mut other_known = false;
    for field in view.fields() {
        match field.number() {
            CUSTOM_REFERENCE_TYPE_FIELD => {},
            CUSTOM_REFERENCE_UUID_FIELD => {
                if uuid_payload.is_some() || field.wire_type() != 2 {
                    return Err(Error::InvalidSource { path });
                }
                field
                    .validate_canonical_framing()
                    .map_err(|_| Error::InvalidSource { path })?;
                uuid_payload = Some(field.payload());
            },
            number if number <= 45 => other_known = true,
            _ => {},
        }
    }
    if format_type.is_none() && uuid_payload.is_none() {
        return Ok(None);
    }
    let format_type = u32::try_from(format_type.ok_or(Error::InvalidSource { path })?)
        .map_err(|_| Error::InvalidSource { path })?;
    if !matches!(
        format_type,
        CUSTOM_NUMBER_FORMAT_TYPE | CUSTOM_TEXT_FORMAT_TYPE | CUSTOM_DATE_TIME_FORMAT_TYPE
    ) {
        if uuid_payload.is_none() {
            return Ok(None);
        }
        return Err(Error::UnsupportedDependency { path });
    }
    if other_known {
        return Err(Error::UnsupportedDependency { path });
    }
    let uuid = parse_custom_uuid(
        uuid_payload.ok_or(Error::InvalidSource { path })?,
        budget,
        path,
    )?;
    Ok(Some(CustomReference { format_type, uuid }))
}

fn validate_custom_legacy_routes(
    source: &Package,
    route: CustomRegistryRoute,
    registry_identifier: u64,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let document_archive = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .ok_or(Error::UnsupportedSource)?;
    let document_object = native::unique_object(document_archive, 1, path)?;
    let document_index =
        native::unique_message_index(document_object, DOCUMENT_MESSAGE_TYPE, path)?;
    let root = &document_object.messages[document_index].data;
    let root_view = custom_wire_view(root, budget, 0, path)?;
    for field in root_view
        .fields()
        .filter(|field| field.number() == DOCUMENT_LEGACY_SUPER_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource { path });
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource { path })?;
        let legacy = custom_wire_view(field.payload(), budget, 1, path)?;
        let mut native_references = 0usize;
        for nested in legacy.fields() {
            match nested.number() {
                7 => return Err(Error::UnsupportedDependency { path }),
                LEGACY_CUSTOM_REGISTRY_REFERENCE_FIELD => {
                    if !matches!(route, CustomRegistryRoute::LegacySuperField12) {
                        return Err(Error::UnsupportedDependency { path });
                    }
                    let nested_identifier = parse_custom_registry_reference(nested, path)?;
                    if nested_identifier != registry_identifier {
                        return Err(Error::InvalidSource { path });
                    }
                    native_references = native_references
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                },
                _ => {},
            }
        }
        if matches!(route, CustomRegistryRoute::LegacySuperField12) && native_references != 1 {
            return Err(Error::InvalidSource { path });
        }
    }
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != 6_001 {
                    continue;
                }
                let (model, report) = storage_codec::decode_table_model_with_report(
                    &message.data,
                    budget.residual_storage_options(&message.data),
                )
                .map_err(|error| native::map_storage_error(error, path))?;
                native::charge_storage_report(budget, report, path)?;
                let (store, report) = storage_codec::decode_data_store_with_report(
                    model.base_data_store(),
                    budget.residual_storage_options(model.base_data_store()),
                )
                .map_err(|error| native::map_storage_error(error, path))?;
                native::charge_storage_report(budget, report, path)?;
                if store.deprecated_custom_format_table().is_some() {
                    return Err(Error::UnsupportedDependency { path });
                }
            }
        }
    }
    Ok(())
}

fn decode_display_payload(
    family: DisplayFormatFamily,
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<NativeDisplayValue, DisplayReadError> {
    let options = native::control_codec_options(source.len(), budget);
    match family {
        DisplayFormatFamily::Number => {
            let (snapshot, report) =
                match number_codec::decode_number_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_decode_failure(family, source, budget, path, error));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let value = NativeDisplayValue::from_parts(
                family,
                snapshot.decimal_places(),
                snapshot.negative_style(),
                snapshot.show_thousands_separator(),
                None,
                None,
                path,
            )?;
            Ok(value)
        },
        DisplayFormatFamily::Currency => {
            let (snapshot, report) =
                match currency_codec::decode_currency_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_decode_failure(family, source, budget, path, error));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let decimal_places = snapshot
                .decimal_places()
                .ok_or(Error::InvalidSource { path })?;
            let currency_code = snapshot
                .currency_code()
                .ok_or(Error::InvalidSource { path })?;
            let negative_style = snapshot
                .negative_style()
                .ok_or(Error::InvalidSource { path })?;
            let show_thousands_separator = snapshot
                .show_thousands_separator()
                .ok_or(Error::InvalidSource { path })?;
            let use_accounting_style = snapshot
                .use_accounting_style()
                .ok_or(Error::InvalidSource { path })?;
            let value = NativeDisplayValue::from_parts(
                family,
                decimal_places,
                negative_style,
                show_thousands_separator,
                Some(currency_code),
                Some(use_accounting_style),
                path,
            )?;
            Ok(value)
        },
        DisplayFormatFamily::Percentage => {
            let (snapshot, report) =
                match percentage_codec::decode_percentage_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_decode_failure(family, source, budget, path, error));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let value = NativeDisplayValue::from_parts(
                family,
                snapshot.decimal_places(),
                snapshot.negative_style(),
                snapshot.show_thousands_separator(),
                None,
                None,
                path,
            )?;
            Ok(value)
        },
        DisplayFormatFamily::Scientific => {
            let (snapshot, report) =
                match scientific_codec::decode_scientific_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_decode_failure(family, source, budget, path, error));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let value = NativeDisplayValue::from_parts(
                family,
                snapshot.decimal_places(),
                snapshot.negative_style(),
                snapshot.show_thousands_separator(),
                None,
                None,
                path,
            )?;
            Ok(value)
        },
        DisplayFormatFamily::Fraction => {
            let (snapshot, report) =
                match fraction_codec::decode_fraction_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_decode_failure(family, source, budget, path, error));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            Ok(NativeDisplayValue::Fraction(fraction_from_native_accuracy(
                snapshot.fraction_accuracy(),
                path,
            )?))
        },
        DisplayFormatFamily::DateTime => {
            let (snapshot, report) =
                match date_time_codec::decode_date_time_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_date_time_decode_failure(
                            source, budget, path, error,
                        ));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let pattern = snapshot.date_time_format();
            let value = DateTime::new(pattern).map_err(|_| Error::InvalidSource { path })?;
            Ok(NativeDisplayValue::DateTime(value))
        },
        DisplayFormatFamily::Duration => {
            let (snapshot, report) =
                match duration_codec::decode_duration_format_with_report(source, options) {
                    Ok(value) => value,
                    Err(error) => {
                        return Err(classify_duration_decode_failure(
                            source, budget, path, error,
                        ));
                    },
                };
            native::charge_control_decode_report(budget, report, path)?;
            let value = duration_from_native(snapshot, path)?;
            Ok(NativeDisplayValue::Duration(value))
        },
        DisplayFormatFamily::Text => {
            let (snapshot, report) = text_codec::decode_text_format_with_report(source, options)
                .map_err(|error| native::map_control_error(error, path))?;
            native::charge_control_decode_report(budget, report, path)?;
            if snapshot.format_type() != TEXT_FORMAT_TYPE {
                return Err(Error::InvalidSource { path }.into());
            }
            Ok(NativeDisplayValue::Text(Text))
        },
    }
}

fn classify_decode_failure(
    family: DisplayFormatFamily,
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
    error: control_codec::DecodeError,
) -> DisplayReadError {
    // The family-specific strict decoder intentionally rejects the sibling
    // discriminator. Probe with the broader strict projection only after a
    // failure, so successful Number/Percentage/Scientific/Fraction reads
    // retain their existing resource accounting. A malformed payload remains
    // a native error.
    let Ok((broad, report)) = control_codec::decode_control_format_with_report(
        source,
        native::control_codec_options(source.len(), budget),
    ) else {
        return DisplayReadError::Native(native::map_control_error(error, path));
    };
    if let Err(native_error) = native::charge_control_decode_report(budget, report, path) {
        return DisplayReadError::Native(native_error);
    }
    if broad.format_type() != family.native_type() {
        DisplayReadError::WrongFormatFamily
    } else {
        DisplayReadError::Native(native::map_control_error(error, path))
    }
}

fn classify_date_time_decode_failure(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
    error: date_time_codec::DecodeError,
) -> DisplayReadError {
    let Ok((broad, report)) = control_codec::decode_control_format_with_report(
        source,
        native::control_codec_options(source.len(), budget),
    ) else {
        return DisplayReadError::Native(native::map_control_error(error, path));
    };
    if let Err(native_error) = native::charge_control_decode_report(budget, report, path) {
        return DisplayReadError::Native(native_error);
    }
    if broad.format_type() != DATE_TIME_FORMAT_TYPE {
        DisplayReadError::WrongFormatFamily
    } else {
        DisplayReadError::Native(native::map_control_error(error, path))
    }
}

fn classify_duration_decode_failure(
    source: &[u8],
    budget: &mut TransactionBudget,
    path: Path,
    error: duration_codec::DecodeError,
) -> DisplayReadError {
    let Ok((broad, report)) = control_codec::decode_control_format_with_report(
        source,
        native::control_codec_options(source.len(), budget),
    ) else {
        return DisplayReadError::Native(native::map_control_error(error, path));
    };
    if let Err(native_error) = native::charge_control_decode_report(budget, report, path) {
        return DisplayReadError::Native(native_error);
    }
    if broad.format_type() != DURATION_FORMAT_TYPE {
        DisplayReadError::WrongFormatFamily
    } else {
        DisplayReadError::Native(native::map_control_error(error, path))
    }
}

fn prepare_display_rewrite(
    family: DisplayFormatFamily,
    source: &[u8],
    value: NativeDisplayValue,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let options = native::control_codec_options(source.len(), budget);
    match family {
        DisplayFormatFamily::Number => {
            let NativeDisplayValue::Number(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = number_codec::prepare_number_format_rewrite(
                source,
                number_codec::NumberFormatWrite::new(
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_number_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Currency => {
            let NativeDisplayValue::Currency(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let code = value.code();
            let prepared = currency_codec::prepare_currency_format_rewrite(
                source,
                currency_codec::CurrencyFormatWrite::new(
                    code.as_str(),
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                    matches!(value.style(), CurrencyStyle::Accounting),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_currency_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Percentage => {
            let NativeDisplayValue::Percentage(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = percentage_codec::prepare_percentage_format_rewrite(
                source,
                percentage_codec::PercentageFormatWrite::new(
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_percentage_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Scientific => {
            let NativeDisplayValue::Scientific(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = scientific_codec::prepare_scientific_format_rewrite(
                source,
                scientific_codec::ScientificFormatWrite::new(u32::from(
                    value.decimal_places().value(),
                )),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_scientific_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Fraction => {
            let NativeDisplayValue::Fraction(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = fraction_codec::prepare_fraction_format_rewrite(
                source,
                fraction_codec::FractionFormatWrite::new(native_fraction_accuracy(
                    value.accuracy(),
                )),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_fraction_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::DateTime => {
            let NativeDisplayValue::DateTime(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = date_time_codec::prepare_date_time_format_rewrite(
                source,
                date_time_codec::DateTimeFormatWrite::new(value.pattern()),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_date_time_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Duration => {
            let NativeDisplayValue::Duration(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = duration_codec::prepare_duration_format_rewrite(
                source,
                duration_format_write(value),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_duration_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Text => {
            let NativeDisplayValue::Text(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = text_codec::prepare_text_format_rewrite(
                source,
                text_codec::TextFormatWrite::new(),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_text_rewrite(prepared, budget, path)
        },
    }
}

fn prepare_display_append(
    family: DisplayFormatFamily,
    value: NativeDisplayValue,
    source_len: usize,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let options = native::control_codec_options(source_len, budget);
    match family {
        DisplayFormatFamily::Number => {
            let NativeDisplayValue::Number(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = number_codec::prepare_number_format_write(
                number_codec::NumberFormatWrite::new(
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_number_append(prepared, budget, path)
        },
        DisplayFormatFamily::Currency => {
            let NativeDisplayValue::Currency(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let code = value.code();
            let prepared = currency_codec::prepare_currency_format_write(
                currency_codec::CurrencyFormatWrite::new(
                    code.as_str(),
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                    matches!(value.style(), CurrencyStyle::Accounting),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_currency_append(prepared, budget, path)
        },
        DisplayFormatFamily::Percentage => {
            let NativeDisplayValue::Percentage(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = percentage_codec::prepare_percentage_format_write(
                percentage_codec::PercentageFormatWrite::new(
                    native_decimal_places(value.decimal_places()),
                    native_negative_style(value.negative_style()),
                    matches!(value.thousands_separator(), ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_percentage_append(prepared, budget, path)
        },
        DisplayFormatFamily::Scientific => {
            let NativeDisplayValue::Scientific(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = scientific_codec::prepare_scientific_format_write(
                scientific_codec::ScientificFormatWrite::new(u32::from(
                    value.decimal_places().value(),
                )),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_scientific_append(prepared, budget, path)
        },
        DisplayFormatFamily::Fraction => {
            let NativeDisplayValue::Fraction(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = fraction_codec::prepare_fraction_format_write(
                fraction_codec::FractionFormatWrite::new(native_fraction_accuracy(
                    value.accuracy(),
                )),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_fraction_append(prepared, budget, path)
        },
        DisplayFormatFamily::DateTime => {
            let NativeDisplayValue::DateTime(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = date_time_codec::prepare_date_time_format_write(
                date_time_codec::DateTimeFormatWrite::new(value.pattern()),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_date_time_append(prepared, budget, path)
        },
        DisplayFormatFamily::Duration => {
            let NativeDisplayValue::Duration(value) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = duration_codec::prepare_duration_format_write(
                duration_format_write(value),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_duration_append(prepared, budget, path)
        },
        DisplayFormatFamily::Text => {
            let NativeDisplayValue::Text(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared =
                text_codec::prepare_text_format_write(text_codec::TextFormatWrite::new(), options)
                    .map_err(|error| native::map_control_error(error, path))?;
            execute_text_append(prepared, budget, path)
        },
    }
}

fn native_decimal_places(value: DecimalPlaces) -> u32 {
    match value {
        DecimalPlaces::Automatic => 253,
        DecimalPlaces::Fixed(value) => u32::from(value.value()),
    }
}

fn duration_format_write(value: Duration) -> duration_codec::DurationFormatWrite {
    let range = value.units().range();
    duration_codec::DurationFormatWrite::from_parts(
        native_duration_style(value.style()),
        native_duration_unit(range.largest()),
        native_duration_unit(range.smallest()),
        value.units().is_automatic(),
    )
}

const fn native_duration_style(value: DurationStyle) -> duration_codec::DurationStyle {
    match value {
        DurationStyle::Colon => duration_codec::DurationStyle::Colon,
        DurationStyle::Abbreviated => duration_codec::DurationStyle::Abbreviated,
        DurationStyle::FullNames => duration_codec::DurationStyle::FullNames,
    }
}

const fn native_duration_unit(value: DurationUnit) -> duration_codec::DurationUnit {
    match value {
        DurationUnit::Weeks => duration_codec::DurationUnit::Weeks,
        DurationUnit::Days => duration_codec::DurationUnit::Days,
        DurationUnit::Hours => duration_codec::DurationUnit::Hours,
        DurationUnit::Minutes => duration_codec::DurationUnit::Minutes,
        DurationUnit::Seconds => duration_codec::DurationUnit::Seconds,
        DurationUnit::Milliseconds => duration_codec::DurationUnit::Milliseconds,
    }
}

fn duration_from_native(
    snapshot: duration_codec::DurationFormatSnapshot<'_>,
    path: Path,
) -> Result<Duration, Error> {
    let style = match snapshot.duration_style() {
        duration_codec::NATIVE_DURATION_STYLE_COLON => DurationStyle::Colon,
        duration_codec::NATIVE_DURATION_STYLE_ABBREVIATED => DurationStyle::Abbreviated,
        duration_codec::NATIVE_DURATION_STYLE_FULL_NAMES => DurationStyle::FullNames,
        _ => return Err(Error::InvalidSource { path }),
    };
    let largest = duration_unit_from_native(snapshot.duration_unit_largest(), path)?;
    let smallest = duration_unit_from_native(snapshot.duration_unit_smallest(), path)?;
    let range = UnitRange::new(largest, smallest).map_err(|_| Error::InvalidSource { path })?;
    let units = if snapshot.use_automatic_duration_units() {
        DurationUnits::Automatic(range)
    } else {
        DurationUnits::Custom(range)
    };
    Ok(Duration::new(style, units))
}

fn duration_unit_from_native(value: u32, path: Path) -> Result<DurationUnit, Error> {
    match value {
        duration_codec::NATIVE_DURATION_UNIT_WEEKS => Ok(DurationUnit::Weeks),
        duration_codec::NATIVE_DURATION_UNIT_DAYS => Ok(DurationUnit::Days),
        duration_codec::NATIVE_DURATION_UNIT_HOURS => Ok(DurationUnit::Hours),
        duration_codec::NATIVE_DURATION_UNIT_MINUTES => Ok(DurationUnit::Minutes),
        duration_codec::NATIVE_DURATION_UNIT_SECONDS => Ok(DurationUnit::Seconds),
        duration_codec::NATIVE_DURATION_UNIT_MILLISECONDS => Ok(DurationUnit::Milliseconds),
        _ => Err(Error::InvalidSource { path }),
    }
}

fn fraction_from_native_accuracy(value: u32, path: Path) -> Result<Fraction, Error> {
    let accuracy = match value {
        fraction_codec::NATIVE_FRACTION_UP_TO_ONE_DIGIT => FractionAccuracy::UpToOneDigit,
        fraction_codec::NATIVE_FRACTION_UP_TO_TWO_DIGITS => FractionAccuracy::UpToTwoDigits,
        fraction_codec::NATIVE_FRACTION_UP_TO_THREE_DIGITS => FractionAccuracy::UpToThreeDigits,
        fraction_codec::NATIVE_FRACTION_HALVES => FractionAccuracy::Halves,
        fraction_codec::NATIVE_FRACTION_QUARTERS => FractionAccuracy::Quarters,
        fraction_codec::NATIVE_FRACTION_EIGHTHS => FractionAccuracy::Eighths,
        fraction_codec::NATIVE_FRACTION_SIXTEENTHS => FractionAccuracy::Sixteenths,
        fraction_codec::NATIVE_FRACTION_TENTHS => FractionAccuracy::Tenths,
        fraction_codec::NATIVE_FRACTION_HUNDREDTHS => FractionAccuracy::Hundredths,
        _ => return Err(Error::InvalidSource { path }),
    };
    Ok(Fraction::new(accuracy))
}

const fn native_fraction_accuracy(value: FractionAccuracy) -> u32 {
    match value {
        FractionAccuracy::UpToOneDigit => fraction_codec::NATIVE_FRACTION_UP_TO_ONE_DIGIT,
        FractionAccuracy::UpToTwoDigits => fraction_codec::NATIVE_FRACTION_UP_TO_TWO_DIGITS,
        FractionAccuracy::UpToThreeDigits => fraction_codec::NATIVE_FRACTION_UP_TO_THREE_DIGITS,
        FractionAccuracy::Halves => fraction_codec::NATIVE_FRACTION_HALVES,
        FractionAccuracy::Quarters => fraction_codec::NATIVE_FRACTION_QUARTERS,
        FractionAccuracy::Eighths => fraction_codec::NATIVE_FRACTION_EIGHTHS,
        FractionAccuracy::Sixteenths => fraction_codec::NATIVE_FRACTION_SIXTEENTHS,
        FractionAccuracy::Tenths => fraction_codec::NATIVE_FRACTION_TENTHS,
        FractionAccuracy::Hundredths => fraction_codec::NATIVE_FRACTION_HUNDREDTHS,
    }
}

const fn native_negative_style(value: NegativeStyle) -> u32 {
    match value {
        NegativeStyle::MinusSign => 0,
        NegativeStyle::Red => 1,
        NegativeStyle::Parentheses => 2,
        NegativeStyle::RedParentheses => 3,
    }
}

fn execute_number_rewrite(
    prepared: number_codec::PreparedNumberFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(number_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_percentage_rewrite(
    prepared: percentage_codec::PreparedPercentageFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(percentage_codec::RewriteExecutionLimits::exact(
            requirements,
        ))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_scientific_rewrite(
    prepared: scientific_codec::PreparedScientificFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(scientific_codec::RewriteExecutionLimits::exact(
            requirements,
        ))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_fraction_rewrite(
    prepared: fraction_codec::PreparedFractionFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(fraction_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_date_time_rewrite(
    prepared: date_time_codec::PreparedDateTimeFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(date_time_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_duration_rewrite(
    prepared: duration_codec::PreparedDurationFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(duration_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_currency_rewrite(
    prepared: currency_codec::PreparedCurrencyFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(currency_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_number_append(
    prepared: number_codec::PreparedNumberFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(number_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_percentage_append(
    prepared: percentage_codec::PreparedPercentageFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(percentage_codec::RewriteExecutionLimits::exact(
            requirements,
        ))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_scientific_append(
    prepared: scientific_codec::PreparedScientificFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(scientific_codec::RewriteExecutionLimits::exact(
            requirements,
        ))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_fraction_append(
    prepared: fraction_codec::PreparedFractionFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(fraction_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_date_time_append(
    prepared: date_time_codec::PreparedDateTimeFormatWrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(date_time_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_duration_append(
    prepared: duration_codec::PreparedDurationFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(duration_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_currency_append(
    prepared: currency_codec::PreparedCurrencyFormatWrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(currency_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_text_append(
    prepared: text_codec::PreparedTextFormatWrite,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(text_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn execute_text_rewrite(
    prepared: text_codec::PreparedTextFormatRewrite<'_>,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let requirements = prepared.execution_requirements();
    native::charge_control_requirements(budget, requirements, path)?;
    let output = prepared
        .execute(text_codec::RewriteExecutionLimits::exact(requirements))
        .map_err(|error| native::map_control_error(error, path))?;
    native::verify_control_report(output.report(), requirements, path)?;
    Ok(output.into_bytes())
}

fn display_read_error_to_write_error(error: DisplayReadError, path: Path) -> Error {
    match error {
        DisplayReadError::WrongFormatFamily => Error::UnsupportedDependency { path },
        DisplayReadError::Native(error) => error,
    }
}

// -------------------------------------------------------------------------
// Document-scoped Custom format owner
// -------------------------------------------------------------------------

// The custom registry is reached only through TN.DocumentArchive field 9.
// The object/message type is intentionally not part of the public contract:
// source-built Numbers files and older producers have used different object
// type tags while retaining this rooted reference.  The format-list entries
// themselves are ordinary TSK.FormatStructArchive values (field 41 is the
// custom UUID).
const CUSTOM_REGISTRY_REFERENCE_FIELD: u32 = 9;
const CUSTOM_REGISTRY_UUID_FIELD: u32 = 1;
const CUSTOM_REGISTRY_FORMAT_FIELD: u32 = 2;
const CUSTOM_REFERENCE_TYPE_FIELD: u32 = 1;
const CUSTOM_REFERENCE_UUID_FIELD: u32 = 41;
const CUSTOM_FORMAT_NAME_FIELD: u32 = 1;
const CUSTOM_FORMAT_PRE_TYPE_FIELD: u32 = 2;
const CUSTOM_FORMAT_DEFAULT_FIELD: u32 = 3;
const CUSTOM_FORMAT_CONDITION_FIELD: u32 = 4;
const CUSTOM_FORMAT_TYPE_FIELD: u32 = 5;
const CUSTOM_PATTERN_TYPE_FIELD: u32 = 1;
const CUSTOM_PATTERN_STRING_FIELD: u32 = 18;
const CUSTOM_CONDITION_TYPE_FIELD: u32 = 1;
const CUSTOM_CONDITION_FLOAT_FIELD: u32 = 2;
const CUSTOM_CONDITION_FORMAT_FIELD: u32 = 3;
const CUSTOM_CONDITION_DOUBLE_FIELD: u32 = 4;
const CUSTOM_NUMBER_FORMAT_TYPE: u32 = 270;
const CUSTOM_TEXT_FORMAT_TYPE: u32 = 271;
const CUSTOM_DATE_TIME_FORMAT_TYPE: u32 = 272;
const CUSTOM_FRACTION_SENTINEL: u32 = (-3_i32) as u32;
const CUSTOM_TEXT_VALUE_TOKEN: char = '\u{e421}';
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
const CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE: u32 = 222;
const FORMAT_LIST_TYPE: i32 = 2;
// TST.TableDataList.ListType::CUSTOM_FORMAT.  This is distinct from the
// document registry's native message type (222).
const CUSTOM_FORMAT_LIST_TYPE: i32 = 6;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
// TN.DocumentArchive.super is field 8, the legacy TSA envelope.  Native
// Numbers has been observed to store the document-scoped custom registry
// reference as field 12 inside that envelope; the compact fixture route uses
// field 9 directly.
const DOCUMENT_LEGACY_SUPER_FIELD: u32 = 8;
const LEGACY_CUSTOM_REGISTRY_REFERENCE_FIELD: u32 = 12;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CustomFormatReadError {
    WrongFormatFamily,
    Native(Error),
}

impl From<Error> for CustomFormatReadError {
    fn from(error: Error) -> Self {
        Self::Native(error)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CustomUuid {
    lower: u64,
    upper: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CustomRegistryEntry {
    uuid: CustomUuid,
    // Existing entries own their semantic projection. Newly appended entries
    // need no second semantic allocation: their canonical payload is already
    // available in the parallel registry vectors and is verified on reopen.
    format: Option<Custom>,
}

#[derive(Debug)]
struct CustomRegistryFacts<'source> {
    source: &'source [u8],
    uuid_payloads: Vec<Cow<'source, [u8]>>,
    format_payloads: Vec<Cow<'source, [u8]>>,
    entries: Vec<CustomRegistryEntry>,
}

#[derive(Debug, Clone, Copy)]
struct CustomRegistryLocation {
    component_index: usize,
    object_identifier: u64,
    message_index: usize,
    message_type: u32,
    route: CustomRegistryRoute,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CustomRegistryRoute {
    DocumentField9,
    LegacySuperField12,
}

#[derive(Debug)]
struct CustomCellGraph {
    tile_component_index: usize,
    tile_identifier: u64,
    tile_message_index: usize,
    tile_payload: Vec<u8>,
    cell_source: Vec<u8>,
    format_component_index: usize,
    format_table_identifier: u64,
    format_message_index: usize,
    format_payload: Vec<u8>,
    format_identifier: Option<u32>,
    explicit_flags: u16,
    cell_kind: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CustomReference {
    format_type: u32,
    uuid: CustomUuid,
}

/// Read one existing cell's document-scoped Custom display format.
pub(super) fn read_custom_format(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<Option<Custom>, CustomFormatReadError> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_custom_format_with_budget(source, target, path, &mut budget)
}

/// Read one Custom display format against a caller-owned transaction ledger.
pub(super) fn read_custom_format_with_budget(
    source: &Package,
    target: CellTarget,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Option<Custom>, CustomFormatReadError> {
    let graph = resolve_custom_cell_graph(source, target, path, false, budget)
        .map_err(CustomFormatReadError::Native)?;
    let Some(format_identifier) = graph.format_identifier else {
        if graph.explicit_flags != 0 {
            return Err(CustomFormatReadError::WrongFormatFamily);
        }
        return Ok(None);
    };
    if graph.explicit_flags == 0 {
        // The format-list edge still has to be structurally valid.  This is
        // an automatic state, not permission to interpret a malformed entry.
        match selected_custom_reference(&graph, format_identifier, path, budget) {
            Ok(reference) => {
                if !custom_reference_matches_cell(
                    reference.format_type,
                    &graph,
                    graph.explicit_flags,
                ) {
                    return Err(CustomFormatReadError::WrongFormatFamily);
                }
            },
            Err(Error::UnsupportedDependency { .. }) => {
                return Err(CustomFormatReadError::WrongFormatFamily);
            },
            Err(error) => return Err(CustomFormatReadError::Native(error)),
        }
        return Ok(None);
    }
    let reference = match selected_custom_reference(&graph, format_identifier, path, budget) {
        Ok(reference) => reference,
        Err(Error::UnsupportedDependency { .. }) => {
            return Err(CustomFormatReadError::WrongFormatFamily);
        },
        Err(error) => return Err(CustomFormatReadError::Native(error)),
    };
    let registry_location =
        locate_custom_registry(source, path, budget).map_err(CustomFormatReadError::Native)?;
    let registry_payload =
        registry_payload(source, registry_location, path).map_err(CustomFormatReadError::Native)?;
    let registry = parse_custom_registry(registry_payload, budget, path)
        .map_err(CustomFormatReadError::Native)?;
    validate_custom_legacy_routes(
        source,
        registry_location.route,
        registry_location.object_identifier,
        path,
        budget,
    )
    .map_err(CustomFormatReadError::Native)?;
    validate_custom_graph_ownership(source, target, &graph, registry_location, path, budget)
        .map_err(CustomFormatReadError::Native)?;
    let custom = registry
        .entries
        .iter()
        .find(|entry| entry.uuid == reference.uuid)
        .ok_or(CustomFormatReadError::Native(Error::InvalidSource { path }))?;
    if !custom_reference_matches_cell(reference.format_type, &graph, graph.explicit_flags) {
        return Err(CustomFormatReadError::WrongFormatFamily);
    }
    let format = custom
        .format
        .as_ref()
        .ok_or(CustomFormatReadError::Native(Error::InvalidSource { path }))?;
    if custom_format_type(format) != reference.format_type {
        return Err(CustomFormatReadError::WrongFormatFamily);
    }
    Ok(Some(format.clone()))
}

/// Rewrite one existing cell's document-scoped Custom display format.
pub(super) fn rewrite_custom_format(
    source: &Package,
    target: CellTarget,
    before: Option<&Custom>,
    after: Option<&Custom>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<native::NativeControlOutput, Error> {
    if target.locked {
        return Err(Error::TableLocked { path });
    }
    if let (Some(before), Some(after)) = (before, after) {
        if custom_format_type(before) != custom_format_type(after) {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    let graph = resolve_custom_cell_graph(source, target, path, true, budget)?;
    let registry_location = locate_custom_registry(source, path, budget)?;
    let registry_payload = registry_payload(source, registry_location, path)?;
    let mut registry = parse_custom_registry(registry_payload, budget, path)?;
    let mut registry_changed = false;
    validate_custom_legacy_routes(
        source,
        registry_location.route,
        registry_location.object_identifier,
        path,
        budget,
    )?;
    validate_custom_graph_ownership(source, target, &graph, registry_location, path, budget)?;

    let current = if let Some(key) = graph.format_identifier {
        if graph.explicit_flags == 0 {
            None
        } else {
            let reference = selected_custom_reference(&graph, key, path, budget)?;
            let entry = registry
                .entries
                .iter()
                .find(|entry| entry.uuid == reference.uuid)
                .ok_or(Error::InvalidSource { path })?;
            let format = entry.format.as_ref().ok_or(Error::InvalidSource { path })?;
            if custom_format_type(format) != reference.format_type
                || !custom_reference_matches_cell(
                    reference.format_type,
                    &graph,
                    graph.explicit_flags,
                )
            {
                return Err(Error::UnsupportedDependency { path });
            }
            Some(format.clone())
        }
    } else {
        if graph.explicit_flags != 0 {
            return Err(Error::UnsupportedDependency { path });
        }
        None
    };
    if current.as_ref() != before {
        return Err(Error::PatchConflict);
    }

    let old_reference = if graph.explicit_flags == 0 {
        None
    } else {
        graph
            .format_identifier
            .map(|key| selected_custom_reference(&graph, key, path, budget))
            .transpose()?
    };
    let old_reference_payload = if old_reference.is_some() {
        graph
            .format_identifier
            .map(|key| selected_custom_reference_payload(&graph, key, path, budget))
            .transpose()?
    } else {
        None
    };
    let desired_reference = if let Some(after) = after {
        let expected_type = custom_format_type(after);
        let uuid = registry
            .entries
            .iter()
            .find(|entry| entry.format.as_ref().is_some_and(|format| format == after))
            .map(|entry| entry.uuid)
            .unwrap_or_else(|| fresh_custom_uuid(&registry.entries));
        if !registry.entries.iter().any(|entry| entry.uuid == uuid) {
            // The source vectors are reserved to their exact input length by
            // the parser. Appending therefore has to be charged and reserved
            // before any new owned payload reaches the allocator.
            budget.charge_allocations(3, path)?;
            registry
                .uuid_payloads
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1, path })?;
            registry
                .format_payloads
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1, path })?;
            registry
                .entries
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1, path })?;
            let (uuid_payload, format_payload) =
                encode_custom_registry_entry(uuid, after, path, budget)?;
            registry.uuid_payloads.push(Cow::Owned(uuid_payload));
            registry.format_payloads.push(Cow::Owned(format_payload));
            registry
                .entries
                .push(CustomRegistryEntry { uuid, format: None });
            registry_changed = true;
        }
        let payload = if let Some(old_payload) = old_reference_payload.as_deref() {
            rewrite_custom_reference(old_payload, expected_type, uuid, path, budget)?
        } else {
            encode_custom_reference(expected_type, uuid, path, budget)?
        };
        Some((uuid, payload))
    } else {
        None
    };

    let desired_payload = desired_reference
        .as_ref()
        .map(|(_, payload)| payload.as_slice());
    let (new_format, new_format_key) = native::mutate_list(
        &graph.format_payload,
        graph.format_identifier,
        desired_payload,
        true,
        budget,
        path,
    )?;
    let replacement_cell = if let Some(key) = new_format_key {
        match after {
            Some(Custom::Text(_)) => {
                rewrite_text_cell_metadata(&graph.cell_source, Some(key), path, budget)?
            },
            Some(Custom::DateTime(_)) => {
                rewrite_date_time_cell_metadata(&graph.cell_source, Some(key), path, budget)?
            },
            Some(Custom::Number(_)) => {
                rewrite_display_cell_metadata(&graph.cell_source, Some(key), path, budget)?
            },
            None => rewrite_custom_cell_clear(&graph.cell_source, path, budget)?,
        }
    } else {
        match before {
            Some(Custom::Text(_)) => {
                rewrite_text_cell_metadata(&graph.cell_source, None, path, budget)?
            },
            Some(Custom::DateTime(_)) => {
                rewrite_date_time_cell_metadata(&graph.cell_source, None, path, budget)?
            },
            Some(Custom::Number(_)) => {
                rewrite_display_cell_metadata(&graph.cell_source, None, path, budget)?
            },
            None => return Err(Error::PatchConflict),
        }
    };

    // The cull decision is made only after the candidate format list exists.
    // Every current direct format list in every component is inspected. A
    // segmented list is an unsupported ownership route, never an opaque list
    // that can be guessed at during UUID cleanup.
    if let Some(old_reference) = old_reference {
        let referenced = custom_uuid_is_referenced(
            source,
            graph.format_component_index,
            graph.format_table_identifier,
            graph.format_message_index,
            &new_format,
            old_reference.uuid,
            budget,
            path,
        )?;
        if !referenced {
            if let Some(index) = registry
                .entries
                .iter()
                .position(|entry| entry.uuid == old_reference.uuid)
            {
                registry.entries.remove(index);
                registry.uuid_payloads.remove(index);
                registry.format_payloads.remove(index);
                registry_changed = true;
            }
        }
    }
    let new_registry = if registry_changed {
        Some(rewrite_custom_registry(&registry, path, budget)?)
    } else {
        None
    };

    charge_display_tile_patch_budget(
        graph.tile_payload.len(),
        replacement_cell.len(),
        budget,
        path,
    )?;
    let patched_tile = popup_native::patch_tile_cell(
        &graph.tile_payload,
        target.position.row(),
        target.position.column(),
        &replacement_cell,
    )
    .map_err(|_| Error::InvalidSource { path })?;

    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let mut component_indices = vec![graph.tile_component_index, graph.format_component_index];
    if registry_changed {
        component_indices.push(registry_location.component_index);
    }
    component_indices.sort_unstable();
    component_indices.dedup();
    let mut archives = component_indices
        .iter()
        .map(|&component_index| {
            let component = source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?;
            budget.charge_archive(
                component.archive(),
                native::archive_serialized_bound(component.archive(), archive_limits, path)?,
                path,
            )?;
            Ok::<_, Error>((component_index, component.archive().clone()))
        })
        .collect::<Result<Vec<_>, Error>>()?;
    popup_native::validate_message_without_object_references(
        source
            .state
            .components
            .catalog()
            .get_index(graph.format_component_index)
            .ok_or(Error::InvalidSource { path })?
            .archive(),
        graph.format_table_identifier,
        graph.format_message_index,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        native::archive_for_mut(&mut archives, graph.format_component_index, path)?,
        graph.format_table_identifier,
        graph.format_message_index,
        new_format,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    popup_native::replace_message_preserving_header(
        native::archive_for_mut(&mut archives, graph.tile_component_index, path)?,
        graph.tile_identifier,
        graph.tile_message_index,
        patched_tile,
        archive_limits,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    if let Some(new_registry) = new_registry {
        popup_native::validate_message_without_object_references(
            source
                .state
                .components
                .catalog()
                .get_index(registry_location.component_index)
                .ok_or(Error::InvalidSource { path })?
                .archive(),
            registry_location.object_identifier,
            registry_location.message_index,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        popup_native::replace_message_preserving_header(
            native::archive_for_mut(&mut archives, registry_location.component_index, path)?,
            registry_location.object_identifier,
            registry_location.message_index,
            new_registry,
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    }

    let mut members = Vec::new();
    members
        .try_reserve_exact(archives.len())
        .map_err(|_| Error::Allocation {
            amount: archives.len(),
            path,
        })?;
    for (component_index, archive) in archives {
        let changed_identifiers = [
            (graph.tile_component_index, graph.tile_identifier),
            (graph.format_component_index, graph.format_table_identifier),
            (
                registry_location.component_index,
                registry_location.object_identifier,
            ),
        ]
        .iter()
        .filter_map(|(owner, identifier)| (*owner == component_index).then_some(*identifier))
        .collect::<Vec<_>>();
        popup_native::verify_archive_object_locality(
            source
                .state
                .components
                .catalog()
                .get_index(component_index)
                .ok_or(Error::InvalidSource { path })?
                .archive(),
            &archive,
            &changed_identifiers,
        )
        .map_err(|_| Error::Verification)?;
        let archive_bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        budget.charge_payload_bytes(archive_bytes.len(), path)?;
        let member_name = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?
            .name()
            .to_owned();
        members.push(native::NativeControlMember {
            member_name,
            archive_bytes,
        });
    }
    Ok(native::NativeControlOutput {
        members,
        component_indices,
        changed_objects: vec![
            (graph.tile_component_index, graph.tile_identifier),
            (graph.format_component_index, graph.format_table_identifier),
            (
                registry_location.component_index,
                registry_location.object_identifier,
            ),
        ],
    })
}

#[cfg(test)]
mod tests {
    use super::{
        CUSTOM_CONDITION_DOUBLE_FIELD, CUSTOM_CONDITION_FORMAT_FIELD, CUSTOM_CONDITION_TYPE_FIELD,
        CUSTOM_DATE_TIME_FORMAT_TYPE, CUSTOM_FRACTION_SENTINEL, CUSTOM_NUMBER_FORMAT_TYPE,
        CUSTOM_PATTERN_STRING_FIELD, CUSTOM_PATTERN_TYPE_FIELD, CUSTOM_REGISTRY_REFERENCE_FIELD,
        CUSTOM_TEXT_FORMAT_TYPE, DOCUMENT_LEGACY_SUPER_FIELD, Package, Path, TransactionBudget,
        append_custom_bytes, append_custom_fixed64, append_custom_varint,
        custom_condition_encoded_len, custom_pattern_encoded_len,
        custom_pattern_index_from_right_last_integer, custom_text_pattern_suffix_width,
        encode_custom_pattern, parse_custom_pattern, parse_optional_varint,
        validate_text_cell_metadata,
    };
    use litchi_iwa_common::wire::WireView;
    use litchi_numbers_wire::BncCell;

    fn encoded_custom_pattern(format_type: u32, pattern: &str) -> Vec<u8> {
        encoded_custom_pattern_with_index(format_type, pattern, 0)
    }

    fn encoded_custom_pattern_with_index(format_type: u32, pattern: &str, index: u64) -> Vec<u8> {
        encoded_custom_pattern_with_metadata(format_type, pattern, index, None)
    }

    fn encoded_custom_pattern_with_metadata(
        format_type: u32,
        pattern: &str,
        index: u64,
        contains_integer: Option<bool>,
    ) -> Vec<u8> {
        let path = Path::Package;
        let mut output = Vec::new();
        append_custom_varint(
            &mut output,
            CUSTOM_PATTERN_TYPE_FIELD,
            u64::from(format_type),
            path,
        )
        .expect("pattern type");
        append_custom_varint(
            &mut output,
            5,
            u64::from(format_type == CUSTOM_NUMBER_FORMAT_TYPE && pattern.contains(',')),
            path,
        )
        .expect("thousands separator");
        append_custom_varint(&mut output, 6, 0, path).expect("field 6");
        append_custom_varint(&mut output, 11, u64::from(CUSTOM_FRACTION_SENTINEL), path)
            .expect("fraction sentinel");
        append_custom_bytes(
            &mut output,
            CUSTOM_PATTERN_STRING_FIELD,
            pattern.as_bytes(),
            path,
        )
        .expect("pattern string");
        append_custom_fixed64(&mut output, 19, 1.0_f64.to_bits(), path).expect("default double");
        append_custom_varint(&mut output, 20, 0, path).expect("field 20");
        for field in [27_u32, 28, 29, 30] {
            append_custom_varint(&mut output, field, 0, path).expect("default field");
        }
        append_custom_varint(&mut output, 31, index, path).expect("index field");
        for field in [34_u32, 35] {
            append_custom_varint(&mut output, field, 0, path).expect("default field");
        }
        append_custom_varint(&mut output, 36, 0, path).expect("field 36");
        append_custom_varint(
            &mut output,
            37,
            u64::from(contains_integer.unwrap_or_else(|| {
                format_type == CUSTOM_NUMBER_FORMAT_TYPE
                    && pattern
                        .chars()
                        .any(|character| matches!(character, '#' | '0'))
            })),
            path,
        )
        .expect("integer placeholder");
        output
    }

    fn test_budget() -> TransactionBudget {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers");
        let package = Package::open(path).expect("Numbers budget fixture");
        TransactionBudget::for_cell_control(&package)
    }

    fn parse_test_pattern(format_type: u32, pattern: &str, index: u64) -> bool {
        let source = encoded_custom_pattern_with_index(format_type, pattern, index);
        let mut budget = test_budget();
        parse_custom_pattern(&source, format_type, &mut budget, Path::Package).is_ok()
    }

    fn encoded_custom_condition(format_type: u32, pattern: &str) -> Vec<u8> {
        let path = Path::Package;
        let nested = encoded_custom_pattern(format_type, pattern);
        let mut output = Vec::new();
        append_custom_varint(&mut output, CUSTOM_CONDITION_TYPE_FIELD, 0, path)
            .expect("condition type");
        append_custom_fixed64(&mut output, CUSTOM_CONDITION_DOUBLE_FIELD, 0, path)
            .expect("condition threshold");
        append_custom_bytes(&mut output, CUSTOM_CONDITION_FORMAT_FIELD, &nested, path)
            .expect("condition format");
        output
    }

    fn hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }

    #[test]
    fn native_marker_zero_text_is_valid_default_metadata() {
        let source = hex("050000000000000000100200050000000c000000");
        let cell = BncCell::parse(&source).expect("native marker-zero Text cell");
        assert_eq!(cell.explicit_format_flags(), 0);
        let metadata = validate_text_cell_metadata(&source, &cell, Path::Package)
            .expect("marker-zero Text metadata");
        assert_eq!(metadata.generic_identifier, None);
        assert_eq!(metadata.text_identifier, Some(12));
    }

    #[test]
    fn native_converted_text_retains_a_generic_number_edge() {
        let source = hex("0503000000008100083002000200000005000000010000000c000000");
        let cell = BncCell::parse(&source).expect("native converted Text cell");
        let metadata = validate_text_cell_metadata(&source, &cell, Path::Package)
            .expect("converted Text metadata");
        assert_eq!(metadata.generic_identifier, Some(1));
        assert_eq!(metadata.text_identifier, Some(12));
    }

    #[test]
    fn reserved_text_metadata_is_rejected_before_publication() {
        let mut source = hex("050000000000800000100200050000000c000000");
        let mut flags = u32::from_le_bytes(source[8..12].try_into().unwrap());
        flags |= 0x0010_0000;
        source[8..12].copy_from_slice(&flags.to_le_bytes());
        source.extend_from_slice(&[0; 4]);
        let cell = BncCell::parse(&source).expect("reserved field remains parseable");
        assert!(validate_text_cell_metadata(&source, &cell, Path::Package).is_err());
    }

    #[test]
    fn custom_registry_root_and_legacy_super_routes_are_distinct() {
        assert_eq!(CUSTOM_REGISTRY_REFERENCE_FIELD, 9);
        assert_eq!(DOCUMENT_LEGACY_SUPER_FIELD, 8);
        assert_ne!(CUSTOM_REGISTRY_REFERENCE_FIELD, DOCUMENT_LEGACY_SUPER_FIELD);
    }

    #[test]
    fn custom_pattern_preflight_matches_number_pattern_content() {
        for pattern in ["0", "#0", "0.00", "#,##0"] {
            let expected = encoded_custom_pattern(CUSTOM_NUMBER_FORMAT_TYPE, pattern).len();
            assert_eq!(
                custom_pattern_encoded_len(CUSTOM_NUMBER_FORMAT_TYPE, pattern).unwrap(),
                expected,
                "pattern {pattern:?} must have an exact preflight length",
            );
        }
    }

    #[test]
    fn native_text_pattern_index_uses_utf16_units() {
        assert_eq!(
            custom_pattern_index_from_right_last_integer(
                CUSTOM_TEXT_FORMAT_TYPE,
                "Native [\u{e421}]",
            )
            .unwrap(),
            9,
        );
        assert_eq!(
            custom_pattern_index_from_right_last_integer(
                CUSTOM_TEXT_FORMAT_TYPE,
                "Rust <\u{e421}>",
            )
            .unwrap(),
            7,
        );
        assert_eq!(
            custom_pattern_index_from_right_last_integer(
                CUSTOM_TEXT_FORMAT_TYPE,
                "😀Native [\u{e421}]",
            )
            .unwrap(),
            11,
        );
        assert_eq!(
            custom_pattern_index_from_right_last_integer(CUSTOM_NUMBER_FORMAT_TYPE, "#,##0")
                .unwrap(),
            0,
        );
        assert_eq!(
            custom_pattern_index_from_right_last_integer(CUSTOM_NUMBER_FORMAT_TYPE, "(#,###)")
                .unwrap(),
            1,
        );
        assert_eq!(
            custom_text_pattern_suffix_width("😀Prefix \u{e421} Suffix"),
            Some(7),
        );
    }

    #[test]
    fn native_text_pattern_cache_profiles_are_parsed_strictly() {
        assert!(parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            "Native [\u{e421}]",
            0
        ));
        assert!(parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            "Native [\u{e421}]",
            9
        ));
        assert!(parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            "😀Native [\u{e421}]",
            11,
        ));
        assert!(!parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            "Native [\u{e421}]",
            8,
        ));
        assert!(!parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            "😀Native [\u{e421}]",
            12,
        ));
        let source_built_text = "😀Prefix \u{e421} Suffix";
        assert!(parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            source_built_text,
            custom_text_pattern_suffix_width(source_built_text).unwrap(),
        ));
        assert!(!parse_test_pattern(
            CUSTOM_TEXT_FORMAT_TYPE,
            source_built_text,
            8,
        ));
        assert!(parse_test_pattern(CUSTOM_NUMBER_FORMAT_TYPE, "(#,###)", 1,));
        assert!(!parse_test_pattern(CUSTOM_NUMBER_FORMAT_TYPE, "#,##0", 1));
        assert!(!parse_test_pattern(
            CUSTOM_DATE_TIME_FORMAT_TYPE,
            "yyyy-MM-dd",
            1,
        ));
        let source_built_date_time = encoded_custom_pattern_with_metadata(
            CUSTOM_DATE_TIME_FORMAT_TYPE,
            "yyyy-MM-dd",
            0,
            Some(true),
        );
        let mut budget = test_budget();
        assert!(
            parse_custom_pattern(
                &source_built_date_time,
                CUSTOM_DATE_TIME_FORMAT_TYPE,
                &mut budget,
                Path::Package,
            )
            .is_ok()
        );
    }

    #[test]
    fn native_text_pattern_cache_rejects_duplicate_and_noncanonical_index() {
        let mut duplicate =
            encoded_custom_pattern_with_index(CUSTOM_TEXT_FORMAT_TYPE, "Native [\u{e421}]", 9);
        append_custom_varint(&mut duplicate, 31, 9, Path::Package).expect("duplicate index");
        let mut budget = test_budget();
        assert!(
            parse_custom_pattern(
                &duplicate,
                CUSTOM_TEXT_FORMAT_TYPE,
                &mut budget,
                Path::Package,
            )
            .is_err()
        );

        let noncanonical = WireView::parse(&[0xf8, 0x01, 0x80, 0x00]).expect("wire view");
        assert!(parse_optional_varint(&noncanonical, 31, Path::Package).is_err());
    }

    #[test]
    fn native_text_pattern_encoder_preflights_index_varint_boundaries() {
        for length in [127_usize, 129] {
            let pattern = "x".repeat(length);
            let mut budget = test_budget();
            let encoded = encode_custom_pattern(
                CUSTOM_TEXT_FORMAT_TYPE,
                &pattern,
                Path::Package,
                &mut budget,
            )
            .expect("encode native Text pattern");
            assert_eq!(
                encoded.len(),
                custom_pattern_encoded_len(CUSTOM_TEXT_FORMAT_TYPE, &pattern)
                    .expect("preflight native Text pattern"),
            );
            let expected_index = u64::try_from(length - 1).expect("index");
            let mut parse_budget = test_budget();
            assert!(
                parse_custom_pattern(
                    &encoded,
                    CUSTOM_TEXT_FORMAT_TYPE,
                    &mut parse_budget,
                    Path::Package,
                )
                .is_ok()
            );
            assert_eq!(
                custom_pattern_index_from_right_last_integer(CUSTOM_TEXT_FORMAT_TYPE, &pattern,)
                    .expect("index helper"),
                expected_index,
            );
        }
    }

    #[test]
    fn custom_condition_preflight_matches_comma_free_and_grouped_patterns() {
        for pattern in ["0", "#0", "#,##0"] {
            let expected = encoded_custom_condition(CUSTOM_NUMBER_FORMAT_TYPE, pattern).len();
            assert_eq!(
                custom_condition_encoded_len(CUSTOM_NUMBER_FORMAT_TYPE, pattern).unwrap(),
                expected,
                "condition pattern {pattern:?} must have an exact preflight length",
            );
        }
    }
}
