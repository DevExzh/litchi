//! One cell's complete display format as the neutral `DataFormat` sum.
//!
//! The per-family readers and writers of
//! [`display`](crate::cell_format::display) and
//! [`pop_up_menu`](crate::cell_format::pop_up_menu) each admit one family
//! and report any other as a wrong family. This module dispatches over
//! them, so a leaf can offer one `edit_*_table_cell_format` family that
//! takes a `DataFormat` instead of
//! one family per format (record 0363 §3.2, record 0379).
//!
//! A rewrite changes one family at a time: Automatic to a format, a format
//! to Automatic, or a format to another value of the same family. A change
//! between two families is two rewrites, the first to Automatic; the leaf
//! publishes the intermediate candidate privately and runs the second rewrite
//! on it.

use litchi_iwa_common::table::cell::data_format::{CellControl, DataFormat};

use super::control::NativeControlOutput;
use super::display::{self, CustomFormatReadError};
use super::{Error, Path, Refusal, Target, TransactionBudget, pop_up_menu};
use crate::TableSource;

/// The family a [`DataFormat`] belongs to; a rewrite stays inside one.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Family {
    /// Decimal Number.
    Number,
    /// Currency.
    Currency,
    /// Percentage.
    Percentage,
    /// Scientific notation.
    Scientific,
    /// Fraction.
    Fraction,
    /// Date and time.
    DateTime,
    /// Duration.
    Duration,
    /// Literal text.
    Text,
    /// Document-scoped custom format.
    Custom,
    /// Numeral system (base-N).
    NumeralSystem,
    /// Checkbox, Star Rating, Slider or Stepper.
    ScalarControl,
    /// Pop-Up Menu.
    PopUpMenu,
}

/// The family of a format, or `None` for Automatic.
#[must_use]
pub const fn family(format: &DataFormat) -> Option<Family> {
    Some(match format {
        DataFormat::Automatic => return None,
        DataFormat::Number(_) => Family::Number,
        DataFormat::Currency(_) => Family::Currency,
        DataFormat::Percentage(_) => Family::Percentage,
        DataFormat::Scientific(_) => Family::Scientific,
        DataFormat::Fraction(_) => Family::Fraction,
        DataFormat::DateTime(_) => Family::DateTime,
        DataFormat::Duration(_) => Family::Duration,
        DataFormat::Text(_) => Family::Text,
        DataFormat::Custom(_) => Family::Custom,
        DataFormat::NumeralSystem(_) => Family::NumeralSystem,
        DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_) => Family::ScalarControl,
        DataFormat::PopUpMenu(_) => Family::PopUpMenu,
    })
}

/// Whether the engine can rewrite this family without a metadata
/// transition: the ten display families, numeral systems included since
/// record 0370k. Controls and Pop-Up Menus also rewrite the package metadata
/// registry.
#[must_use]
pub const fn is_display_family(family: Family) -> bool {
    !matches!(family, Family::ScalarControl | Family::PopUpMenu)
}

/// Read the selected stored cell's display format.
///
/// # Errors
///
/// The engine's typed failures; [`Refusal::UnsupportedDependency`] when the
/// cell holds a format no family reader admits (a native shape outside the
/// engines).
pub fn read(source: &dyn TableSource, target: Target, path: Path) -> Result<DataFormat, Error> {
    let mut budget = TransactionBudget::for_cell_control(source);
    read_with_budget(source, target, path, &mut budget)
}

/// Read the selected cell's complete format against a caller-owned
/// transaction ledger.  Candidate verification uses this entry point so its
/// semantic readback cannot create a fresh budget after publication work has
/// already consumed the operation's original one.
pub fn read_with_budget(
    source: &dyn TableSource,
    target: Target,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<DataFormat, Error> {
    if let Some(control) = pop_up_menu::read_cell_control(source, target)? {
        return Ok(match control {
            CellControl::Checkbox(value) => DataFormat::Checkbox(value),
            CellControl::StarRating(value) => DataFormat::StarRating(value),
            CellControl::Slider(value) => DataFormat::Slider(value),
            CellControl::Stepper(value) => DataFormat::Stepper(value),
            CellControl::PopUpMenu(value) => DataFormat::PopUpMenu(value),
        });
    }
    // One validation pass for the ten display families (record 0370k); a
    // cell outside them is a custom format or unsupported.
    if let Some(format) = display::read_any_display_format(source, target, path, budget)? {
        return Ok(format);
    }
    match display::read_custom_format_with_budget(source, target, path, budget) {
        Ok(Some(value)) => Ok(DataFormat::Custom(value)),
        Ok(None) => Ok(DataFormat::Automatic),
        Err(CustomFormatReadError::WrongFormatFamily) => {
            Err(Error::Refused(Refusal::UnsupportedDependency { path }))
        },
        Err(CustomFormatReadError::Native(error)) => Err(error),
    }
}

/// Rewrite one display family of the selected stored cell.
///
/// `before` and `after` must not belong to two different families, and the
/// family must be a display family ([`is_display_family`]).
///
/// # Errors
///
/// [`Refusal::UnsupportedDependency`] for a cross-family change or a family
/// outside the display engines; the engine's typed failures otherwise.
pub fn rewrite(
    source: &dyn TableSource,
    target: Target,
    before: &DataFormat,
    after: &DataFormat,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeControlOutput, Error> {
    let unsupported = Error::Refused(Refusal::UnsupportedDependency { path });
    let selected = match (family(before), family(after)) {
        (Some(left), Some(right)) if left != right => return Err(unsupported),
        (Some(family), _) | (None, Some(family)) => family,
        (None, None) => return Err(unsupported),
    };
    if !is_display_family(selected) {
        return Err(unsupported);
    }
    macro_rules! rewrite_with {
        ($writer:ident, $variant:ident) => {{
            let before = match before {
                DataFormat::$variant(value) => Some(value),
                _ => None,
            };
            let after = match after {
                DataFormat::$variant(value) => Some(value),
                _ => None,
            };
            display::$writer(source, target, before, after, path, budget)
        }};
    }
    match selected {
        Family::Number => rewrite_with!(rewrite_number_format, Number),
        Family::Currency => rewrite_with!(rewrite_currency_format, Currency),
        Family::Percentage => rewrite_with!(rewrite_percentage_format, Percentage),
        Family::Scientific => rewrite_with!(rewrite_scientific_format, Scientific),
        Family::Fraction => rewrite_with!(rewrite_fraction_format, Fraction),
        Family::DateTime => rewrite_with!(rewrite_date_time_format, DateTime),
        Family::Duration => rewrite_with!(rewrite_duration_format, Duration),
        Family::Text => rewrite_with!(rewrite_text_format, Text),
        Family::Custom => rewrite_with!(rewrite_custom_format, Custom),
        Family::NumeralSystem => rewrite_with!(rewrite_numeral_system_format, NumeralSystem),
        Family::ScalarControl | Family::PopUpMenu => Err(unsupported),
    }
}

/// The number of members a display rewrite of `family` may return.
#[must_use]
pub const fn maximum_members(family: Family) -> usize {
    match family {
        // A scalar control rewrites the tile, the format list and the
        // control list.
        Family::Custom | Family::ScalarControl => 3,
        _ => 2,
    }
}

/// The scalar control a format holds, when it is one (Checkbox, Star
/// Rating, Slider or Stepper).
#[must_use]
pub fn scalar_control(format: &DataFormat) -> Option<CellControl> {
    Some(match format {
        DataFormat::Checkbox(value) => CellControl::Checkbox(*value),
        DataFormat::StarRating(value) => CellControl::StarRating(*value),
        DataFormat::Slider(value) => CellControl::Slider(value.clone()),
        DataFormat::Stepper(value) => CellControl::Stepper(value.clone()),
        _ => return None,
    })
}

/// Rewrite the selected stored cell within one family: the display
/// families through [`rewrite`], the scalar controls through
/// [`control::rewrite_scalar_control`](super::control::rewrite_scalar_control)
/// (record 0412). Pop-Up Menus are refused here; their writes also allocate
/// a menu model and rewrite the package metadata registry.
///
/// # Errors
///
/// As [`rewrite`]; [`Refusal::UnsupportedDependency`] for a Pop-Up Menu or
/// a change between two families.
pub fn rewrite_family(
    source: &dyn TableSource,
    target: Target,
    before: &DataFormat,
    after: &DataFormat,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeControlOutput, Error> {
    let scalar = |format: &DataFormat| family(format) == Some(Family::ScalarControl);
    if scalar(before) || scalar(after) {
        if (!scalar(before) && family(before).is_some())
            || (!scalar(after) && family(after).is_some())
        {
            return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
        }
        return super::control::rewrite_scalar_control(
            source,
            target,
            scalar_control(before).as_ref(),
            scalar_control(after).as_ref(),
            path,
            budget,
        );
    }
    rewrite(source, target, before, after, path, budget)
}
