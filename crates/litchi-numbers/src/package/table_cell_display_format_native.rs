//! Private native owner for focused decimal display-format transactions.
//!
//! Number, Percentage, Scientific, and Fraction cells use the same BNC decimal
//! cell kind and the same format-list graph. This module keeps that graph
//! surgery in one route while retaining a typed family boundary at every
//! codec and semantic conversion.
//! It intentionally exposes no native identifiers or arbitrary format-type
//! input to the package API.

use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_currency_format_codec as currency_codec,
    numbers_table_cell_fraction_format_codec as fraction_codec,
    numbers_table_cell_number_format_codec as number_codec,
    numbers_table_cell_percentage_format_codec as percentage_codec,
    numbers_table_cell_scientific_format_codec as scientific_codec,
    numbers_table_cell_storage_codec as storage_codec,
};
use litchi_numbers_wire::{BncCell, NumericCellType};

use super::table_cell_control_native as native;
use super::{
    Package, table_cell_pop_up_menu as popup,
    table_cell_pop_up_menu::{CellTarget, Error, Path, TransactionBudget},
    table_cell_pop_up_menu_native as popup_native,
};
use crate::cell::data_format::currency::{Currency, CurrencyCode, CurrencyStyle};
use crate::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, Fraction, FractionAccuracy, NegativeStyle, Number,
    Percentage, Scientific, ThousandsSeparator,
};

const NUMBER_FORMAT_TYPE: u32 = number_codec::NATIVE_NUMBER_FORMAT_TYPE;
const CURRENCY_FORMAT_TYPE: u32 = currency_codec::NATIVE_CURRENCY_FORMAT_TYPE;
const PERCENTAGE_FORMAT_TYPE: u32 = percentage_codec::NATIVE_PERCENTAGE_FORMAT_TYPE;
const SCIENTIFIC_FORMAT_TYPE: u32 = scientific_codec::NATIVE_SCIENTIFIC_FORMAT_TYPE;
const FRACTION_FORMAT_TYPE: u32 = fraction_codec::NATIVE_FRACTION_FORMAT_TYPE;

/// The only display families admitted by the focused decimal owner.
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
}

impl DisplayFormatFamily {
    const fn native_type(self) -> u32 {
        match self {
            Self::Number => NUMBER_FORMAT_TYPE,
            Self::Currency => CURRENCY_FORMAT_TYPE,
            Self::Percentage => PERCENTAGE_FORMAT_TYPE,
            Self::Scientific => SCIENTIFIC_FORMAT_TYPE,
            Self::Fraction => FRACTION_FORMAT_TYPE,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NativeDisplayValue {
    Number(Number),
    Currency(Currency),
    Percentage(Percentage),
    Scientific(Scientific),
    Fraction(Fraction),
}

impl NativeDisplayValue {
    const fn family(self) -> DisplayFormatFamily {
        match self {
            Self::Number(_) => DisplayFormatFamily::Number,
            Self::Currency(_) => DisplayFormatFamily::Currency,
            Self::Percentage(_) => DisplayFormatFamily::Percentage,
            Self::Scientific(_) => DisplayFormatFamily::Scientific,
            Self::Fraction(_) => DisplayFormatFamily::Fraction,
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
        | Err(DisplayReadError::WrongFormatFamily) => {
            Err(FractionFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(FractionFormatReadError::Native(error)),
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
    let format_identifier = cell.format_identifier();
    if cell.control_cell_spec_identifier().is_some() {
        return Err(DisplayReadError::WrongFormatFamily);
    }
    let explicit_flags = cell.explicit_format_flags();
    let secondary_identifier = cell.secondary_format_identifier();
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
        (_, Some(_)) => return Err(DisplayReadError::WrongFormatFamily),
        (_, None) if format_identifier.is_some() || explicit_flags != 0 => {
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
    if before.is_some_and(|value| value.family() != family)
        || after.is_some_and(|value| value.family() != family)
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
    let old_format = cell.format_identifier();
    let old_secondary = cell.secondary_format_identifier();
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
        (_, None, None)
            if explicit_flags == 0
                && old_secondary.is_none()
                && (family != DisplayFormatFamily::Currency
                    || currency_cell_type_matches_value(&cell, false)) => {},
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

    if family == DisplayFormatFamily::Currency {
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
    if family == DisplayFormatFamily::Currency && after.is_none() {
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
    let replacement_cell = if after.is_some() {
        let key = new_format_key.ok_or(Error::InvalidSource { path })?;
        if family == DisplayFormatFamily::Currency {
            rewrite_currency_cell_metadata(cell_source, Some(key), path, budget)?
        } else {
            rewrite_display_cell_metadata(cell_source, Some(key), path, budget)?
        }
    } else {
        if family == DisplayFormatFamily::Currency {
            rewrite_currency_cell_metadata(cell_source, None, path, budget)?
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
    }
}

fn native_decimal_places(value: DecimalPlaces) -> u32 {
    match value {
        DecimalPlaces::Automatic => 253,
        DecimalPlaces::Fixed(value) => u32::from(value.value()),
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

fn display_read_error_to_write_error(error: DisplayReadError, path: Path) -> Error {
    match error {
        DisplayReadError::WrongFormatFamily => Error::UnsupportedDependency { path },
        DisplayReadError::Native(error) => error,
    }
}
