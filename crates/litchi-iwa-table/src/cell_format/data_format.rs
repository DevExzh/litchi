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
use super::display::{
    self, CurrencyFormatReadError, CustomFormatReadError, DateTimeFormatReadError,
    DurationFormatReadError, FractionFormatReadError, NumberFormatReadError,
    PercentageFormatReadError, ScientificFormatReadError, TextFormatReadError,
};
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
/// transition: the nine display families. Controls and Pop-Up Menus also
/// rewrite the package metadata registry, and numeral systems have no
/// engine yet.
#[must_use]
pub const fn is_display_family(family: Family) -> bool {
    !matches!(
        family,
        Family::NumeralSystem | Family::ScalarControl | Family::PopUpMenu
    )
}

/// One family read: `Some(None)` reports an automatic cell, `None` a cell of
/// another family.
fn family_read<T, E>(
    result: Result<Option<T>, E>,
    wrong: impl Fn(&E) -> bool,
    native: impl Fn(E) -> Error,
) -> Result<Option<Option<T>>, Error> {
    match result {
        Ok(value) => Ok(Some(value)),
        Err(error) if wrong(&error) => Ok(None),
        Err(error) => Err(native(error)),
    }
}

macro_rules! read_family {
    ($reader:ident, $error:ident, $variant:ident, $source:expr, $target:expr, $path:expr, $automatic:ident) => {
        match family_read(
            display::$reader($source, $target, $path),
            |error| matches!(error, $error::WrongFormatFamily),
            |error| match error {
                $error::WrongFormatFamily => Error::Verification,
                $error::Native(error) => error,
            },
        )? {
            Some(Some(value)) => return Ok(DataFormat::$variant(value)),
            Some(None) => $automatic = true,
            None => {},
        }
    };
}

/// Read the selected stored cell's display format.
///
/// # Errors
///
/// The engine's typed failures; [`Refusal::UnsupportedDependency`] when the
/// cell holds a format no family reader admits (a numeral system, or a
/// native shape outside the engines).
pub fn read(source: &dyn TableSource, target: Target, path: Path) -> Result<DataFormat, Error> {
    if let Some(control) = pop_up_menu::read_cell_control(source, target)? {
        return Ok(match control {
            CellControl::Checkbox(value) => DataFormat::Checkbox(value),
            CellControl::StarRating(value) => DataFormat::StarRating(value),
            CellControl::Slider(value) => DataFormat::Slider(value),
            CellControl::Stepper(value) => DataFormat::Stepper(value),
            CellControl::PopUpMenu(value) => DataFormat::PopUpMenu(value),
        });
    }
    let mut automatic = false;
    read_family!(
        read_number_format,
        NumberFormatReadError,
        Number,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_currency_format,
        CurrencyFormatReadError,
        Currency,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_percentage_format,
        PercentageFormatReadError,
        Percentage,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_scientific_format,
        ScientificFormatReadError,
        Scientific,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_fraction_format,
        FractionFormatReadError,
        Fraction,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_date_time_format,
        DateTimeFormatReadError,
        DateTime,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_duration_format,
        DurationFormatReadError,
        Duration,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_text_format,
        TextFormatReadError,
        Text,
        source,
        target,
        path,
        automatic
    );
    read_family!(
        read_custom_format,
        CustomFormatReadError,
        Custom,
        source,
        target,
        path,
        automatic
    );
    if automatic {
        Ok(DataFormat::Automatic)
    } else {
        Err(Error::Refused(Refusal::UnsupportedDependency { path }))
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
        Family::NumeralSystem | Family::ScalarControl | Family::PopUpMenu => Err(unsupported),
    }
}

/// The number of members a display rewrite of `family` may return.
#[must_use]
pub const fn maximum_members(family: Family) -> usize {
    match family {
        Family::Custom => 3,
        _ => 2,
    }
}
