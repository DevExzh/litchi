//! Typed data-format storage for native table cells.

use super::*;
use crate::numbers::bnc;
use crate::protobuf::tsk::FormatStructArchive;
use crate::table_cell_data_format::{
    TableCellCheckboxFormat, TableCellCurrencyFormat, TableCellDataFormat, TableCellDateTimeFormat,
    TableCellDurationFormat, TableCellFractionFormat, TableCellNumberFormat,
    TableCellNumeralSystemFormat, TableCellNumericControlDisplayFormat, TableCellPercentageFormat,
    TableCellScientificFormat, TableCellSliderFormat, TableCellStarRatingFormat,
    TableCellStepperFormat,
};
#[cfg(test)]
use crate::table_cell_data_format::{
    TableCellCurrencyCode, TableCellCurrencyStyle, TableCellDecimalPlaces,
    TableCellFixedDecimalPlaces, TableCellNegativeNumberStyle, TableCellSliderRange,
    TableCellStepperRange, TableCellThousandsSeparator,
};

mod codec;
mod control;
use codec::*;

pub(super) fn cell_data_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TableCellDataFormat> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(TableCellDataFormat::Automatic);
    };
    let cell = BncCell::parse(&data)?;
    match format_reference(&cell)? {
        CellFormatReference::Automatic { .. } => Ok(TableCellDataFormat::Automatic),
        CellFormatReference::Explicit { identifier, .. } => {
            let resolved = resolve_format_table(package, &location)?;
            let entry = required_format_entry(&resolved, identifier)?;
            let native = entry.entry.format.as_ref().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers format-table entry {identifier} has no format payload"
                ))
            })?;
            let control = cell
                .control_cell_spec_identifier()
                .map(|identifier| control::read_spec(package, &location, identifier))
                .transpose()?;
            match control {
                None => {
                    let format = data_format_from_native(native)?;
                    if matches!(
                        format,
                        TableCellDataFormat::Checkbox(_)
                            | TableCellDataFormat::StarRating(_)
                            | TableCellDataFormat::Slider(_)
                            | TableCellDataFormat::Stepper(_)
                    ) {
                        return Err(Error::InvalidFormat(
                            "Interactive cell format has no control-cell-spec reference".to_owned(),
                        ));
                    }
                    Ok(format)
                },
                Some(control::ControlCellSpecKind::Checkbox) => {
                    if cell.cell_format_kind() != Some(bnc::CHECKBOX_CELL_FORMAT_KIND) {
                        return Err(Error::InvalidFormat(
                            "Checkbox control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    let format = data_format_from_native(native)?;
                    if !matches!(format, TableCellDataFormat::Checkbox(_)) {
                        return Err(Error::InvalidFormat(
                            "Checkbox control references a non-Checkbox format".to_owned(),
                        ));
                    }
                    Ok(format)
                },
                Some(control::ControlCellSpecKind::StarRating) => {
                    if cell.cell_format_kind() != Some(bnc::STAR_RATING_CELL_FORMAT_KIND) {
                        return Err(Error::InvalidFormat(
                            "Star Rating control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    let format = data_format_from_native(native)?;
                    if !matches!(format, TableCellDataFormat::StarRating(_)) {
                        return Err(Error::InvalidFormat(
                            "Star Rating control references a non-Star-Rating format".to_owned(),
                        ));
                    }
                    Ok(format)
                },
                Some(control::ControlCellSpecKind::Slider(range)) => {
                    let display_format = numeric_control_display_from_native(native)?;
                    let expected_kind = match display_format {
                        TableCellNumericControlDisplayFormat::Currency(_) => {
                            bnc::CURRENCY_CELL_FORMAT_KIND
                        },
                        _ => bnc::DECIMAL_CELL_FORMAT_KIND,
                    };
                    if cell.cell_format_kind() != Some(expected_kind) {
                        return Err(Error::InvalidFormat(
                            "Slider control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    Ok(TableCellDataFormat::Slider(TableCellSliderFormat::new(
                        range,
                        display_format,
                    )))
                },
                Some(control::ControlCellSpecKind::Stepper(range)) => {
                    let display_format = numeric_control_display_from_native(native)?;
                    let expected_kind = match display_format {
                        TableCellNumericControlDisplayFormat::Currency(_) => {
                            bnc::CURRENCY_CELL_FORMAT_KIND
                        },
                        _ => bnc::DECIMAL_CELL_FORMAT_KIND,
                    };
                    if cell.cell_format_kind() != Some(expected_kind) {
                        return Err(Error::InvalidFormat(
                            "Stepper control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    Ok(TableCellDataFormat::Stepper(TableCellStepperFormat::new(
                        range,
                        display_format,
                    )))
                },
            }
        },
    }
}

pub(super) fn cell_number_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellNumberFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Number(format) => Ok(Some(format)),
        TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Number data format".to_owned(),
        )),
    }
}

pub(super) fn cell_currency_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellCurrencyFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Currency(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Currency data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_currency_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Currency(_) => reset_cell_data_format(package, table_id, row, column),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Currency format from a non-Currency cell".to_owned(),
        )),
    }
}

pub(super) fn cell_percentage_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellPercentageFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Percentage(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Percentage data format".to_owned(),
        )),
    }
}

pub(super) fn cell_scientific_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellScientificFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Scientific(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Scientific data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_scientific_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Scientific(_) => {
            reset_cell_data_format(package, table_id, row, column)
        },
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Scientific format from a non-Scientific cell".to_owned(),
        )),
    }
}

pub(super) fn cell_fraction_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellFractionFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Fraction(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Fraction data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_fraction_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Fraction(_) => reset_cell_data_format(package, table_id, row, column),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Fraction format from a non-Fraction cell".to_owned(),
        )),
    }
}

pub(super) fn cell_numeral_system_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellNumeralSystemFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::NumeralSystem(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Numeral System data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_numeral_system_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::NumeralSystem(_) => {
            reset_cell_data_format(package, table_id, row, column)
        },
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Numeral System format from a non-Numeral-System cell".to_owned(),
        )),
    }
}

pub(super) fn cell_date_time_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellDateTimeFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::DateTime(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Date & Time data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_date_time_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::DateTime(_) => reset_cell_data_format(package, table_id, row, column),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Date & Time format from a non-Date-Time cell".to_owned(),
        )),
    }
}

pub(super) fn cell_duration_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellDurationFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Duration(format) => Ok(Some(format)),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Duration data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_duration_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Duration(_) => reset_cell_data_format(package, table_id, row, column),
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Duration format from a non-Duration cell".to_owned(),
        )),
    }
}

pub(super) fn cell_checkbox_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellCheckboxFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Checkbox(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Checkbox data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_checkbox_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Checkbox(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Checkbox format from a non-Checkbox cell".to_owned(),
        )),
    }
}

pub(super) fn cell_star_rating_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellStarRatingFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::StarRating(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Star Rating data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_star_rating_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::StarRating(_) => {
            reset_cell_data_format(package, table_id, row, column)
        },
        _ => Err(Error::InvalidFormat(
            "Cannot reset Star Rating format from a non-Star-Rating cell".to_owned(),
        )),
    }
}

pub(super) fn cell_slider_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellSliderFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Slider(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Slider data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_slider_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Slider(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Slider format from a non-Slider cell".to_owned(),
        )),
    }
}

pub(super) fn cell_stepper_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellStepperFormat>> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(None),
        TableCellDataFormat::Stepper(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Stepper data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_stepper_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Stepper(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Stepper format from a non-Stepper cell".to_owned(),
        )),
    }
}

pub(super) fn set_cell_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: TableCellNumberFormat,
) -> Result<()> {
    set_cell_data_format(package, table_id, row, column, &format.into())
}

pub(super) fn set_cell_data_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: &TableCellDataFormat,
) -> Result<()> {
    if format == &TableCellDataFormat::Automatic {
        reset_cell_data_format(package, table_id, row, column)?;
        return Ok(());
    }
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    ensure_current_format_table(package, table_id, row, column)?;
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let existing_data = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?;
    let mut cell = existing_data
        .as_deref()
        .map(BncCell::parse)
        .transpose()?
        .unwrap_or_else(BncCell::minimal);
    let old_reference = format_reference(&cell)?;
    let old_identifier = old_reference.primary_identifier();
    let old_identifiers = old_reference.identifiers();
    let old_control_identifier = cell.control_cell_spec_identifier();
    let cell_format_kind = match format {
        TableCellDataFormat::Currency(_) => bnc::CellDataFormatKind::Currency,
        TableCellDataFormat::DateTime(_) => bnc::CellDataFormatKind::DateTime,
        TableCellDataFormat::Duration(_) => bnc::CellDataFormatKind::Duration,
        TableCellDataFormat::Checkbox(_) => bnc::CellDataFormatKind::Checkbox,
        TableCellDataFormat::StarRating(_) => bnc::CellDataFormatKind::StarRating,
        TableCellDataFormat::Slider(format) => match format.display_format() {
            TableCellNumericControlDisplayFormat::Currency(_) => {
                bnc::CellDataFormatKind::NumericControlCurrency
            },
            _ => bnc::CellDataFormatKind::NumericControlNumberOrPercentage,
        },
        TableCellDataFormat::Stepper(format) => match format.display_format() {
            TableCellNumericControlDisplayFormat::Currency(_) => {
                bnc::CellDataFormatKind::NumericControlCurrency
            },
            _ => bnc::CellDataFormatKind::NumericControlNumberOrPercentage,
        },
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_) => bnc::CellDataFormatKind::NumberOrPercentage,
        TableCellDataFormat::Automatic => unreachable!("handled above"),
    };
    let native = data_format_to_native(format)?;
    let new_control_identifier = match format {
        TableCellDataFormat::Checkbox(_) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Checkbox,
        )?),
        TableCellDataFormat::StarRating(_) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::StarRating,
        )?),
        TableCellDataFormat::Slider(format) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Slider(format.range()),
        )?),
        TableCellDataFormat::Stepper(format) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Stepper(format.range()),
        )?),
        _ => None,
    };
    let resolved = resolve_format_table(package, &location)?;
    let old_entries = old_identifiers
        .into_iter()
        .flatten()
        .map(|identifier| required_format_entry(&resolved, identifier))
        .collect::<Result<Vec<_>>>()?;
    let old_entry = old_identifier
        .and_then(|identifier| {
            old_entries
                .iter()
                .find(|entry| entry.entry.key == identifier)
        })
        .copied();

    if matches!(old_reference, CellFormatReference::Explicit { .. })
        && old_entry.and_then(|entry| entry.entry.format.as_ref()) == Some(&native)
        && old_control_identifier == new_control_identifier
    {
        return Ok(());
    }

    let reusable = resolved
        .entries
        .iter()
        .find(|located| located.entry.format.as_ref() == Some(&native));
    if reusable.is_some_and(|entry| entry.entry.refcount == 0) {
        return Err(Error::InvalidFormat(
            "Numbers format table contains a zero-reference entry".to_owned(),
        ));
    }
    let new_identifier = if let Some(reusable) = reusable {
        if !old_identifiers.contains(&Some(reusable.entry.key)) {
            storage::increment_table_data_list_entry(
                package,
                &location.object_locations,
                &resolved,
                reusable,
                tst::table_data_list::ListType::Format,
            )?;
        }
        reusable.entry.key
    } else {
        append_format_entry(package, &resolved, native)?
    };

    if let TableCellDataFormat::Slider(format) = format
        && cell.cached_scalar()?.is_none()
    {
        cell.set_plain_number(format.range().native_initial_value())?;
    }
    if let TableCellDataFormat::Stepper(format) = format
        && cell.cached_scalar()?.is_none()
    {
        cell.set_plain_number(format.range().native_initial_value())?;
    }
    cell.set_data_format_identifier(new_identifier, cell_format_kind, new_control_identifier)?;
    let cell_count = storage::update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        model::EncodedValue::Raw(cell.encode()),
    )?;
    storage::update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )?;

    for old_entry in old_entries {
        if old_entry.entry.key != new_identifier {
            storage::decrement_table_data_list_entry(
                package,
                &location.object_locations,
                &resolved,
                old_entry,
                tst::table_data_list::ListType::Format,
            )?;
        }
    }
    if let Some(identifier) = old_control_identifier
        && Some(identifier) != new_control_identifier
    {
        control::release_spec(package, &location, identifier)?;
    }
    Ok(())
}

pub(super) fn reset_cell_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Number(_) => reset_cell_data_format(package, table_id, row, column),
        TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Number format from a non-Number cell".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_percentage_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        TableCellDataFormat::Automatic => Ok(false),
        TableCellDataFormat::Percentage(_) => {
            reset_cell_data_format(package, table_id, row, column)
        },
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Currency(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_)
        | TableCellDataFormat::NumeralSystem(_)
        | TableCellDataFormat::DateTime(_)
        | TableCellDataFormat::Duration(_)
        | TableCellDataFormat::Checkbox(_)
        | TableCellDataFormat::StarRating(_)
        | TableCellDataFormat::Slider(_)
        | TableCellDataFormat::Stepper(_) => Err(Error::InvalidFormat(
            "Cannot reset Percentage format from a non-Percentage cell".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_data_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(false);
    };
    let mut cell = BncCell::parse(&data)?;
    let reference = format_reference(&cell)?;
    let CellFormatReference::Explicit { .. } = reference else {
        return Ok(false);
    };
    let resolved = resolve_format_table(package, &location)?;
    let old_entries = reference
        .identifiers()
        .into_iter()
        .flatten()
        .map(|identifier| required_format_entry(&resolved, identifier))
        .collect::<Result<Vec<_>>>()?;
    let old_control_identifier = cell.control_cell_spec_identifier();

    cell.clear_explicit_format();
    let cell_count = storage::update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        model::EncodedValue::Raw(cell.encode()),
    )?;
    storage::update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )?;
    for old_entry in old_entries {
        storage::decrement_table_data_list_entry(
            package,
            &location.object_locations,
            &resolved,
            old_entry,
            tst::table_data_list::ListType::Format,
        )?;
    }
    if let Some(identifier) = old_control_identifier {
        control::release_spec(package, &location, identifier)?;
    }
    Ok(true)
}

#[derive(Clone, Copy)]
enum CellFormatReference {
    Automatic {
        identifier: Option<u32>,
        secondary: Option<u32>,
    },
    Explicit {
        identifier: u32,
        secondary: Option<u32>,
    },
}

impl CellFormatReference {
    const fn primary_identifier(self) -> Option<u32> {
        match self {
            Self::Automatic { identifier, .. } => identifier,
            Self::Explicit { identifier, .. } => Some(identifier),
        }
    }

    const fn identifiers(self) -> [Option<u32>; 2] {
        match self {
            Self::Automatic {
                identifier,
                secondary,
            } => [identifier, secondary],
            Self::Explicit {
                identifier,
                secondary,
            } => [Some(identifier), secondary],
        }
    }
}

fn format_reference(cell: &BncCell) -> Result<CellFormatReference> {
    let explicit = cell.explicit_format_flags();
    let kind = cell.cell_format_kind();
    let identifier = cell.format_identifier();
    let secondary = cell.secondary_format_identifier();
    match (kind, cell.control_cell_spec_identifier()) {
        (Some(bnc::CHECKBOX_CELL_FORMAT_KIND), Some(_))
        | (Some(bnc::DECIMAL_CELL_FORMAT_KIND), _)
        | (Some(bnc::CURRENCY_CELL_FORMAT_KIND), _)
        | (_, None) => {},
        _ => {
            return Err(Error::InvalidFormat(
                "Table cell contains inconsistent interactive-control metadata".to_owned(),
            ));
        },
    }
    match (explicit, kind, identifier) {
        (0, None, None) if secondary.is_none() => Ok(CellFormatReference::Automatic {
            identifier: None,
            secondary: None,
        }),
        (0, Some(bnc::DECIMAL_CELL_FORMAT_KIND), Some(identifier)) if secondary.is_none() => {
            Ok(CellFormatReference::Automatic {
                identifier: Some(identifier),
                secondary: None,
            })
        },
        (0, Some(bnc::CURRENCY_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Automatic {
                identifier: Some(identifier),
                secondary,
            })
        },
        (0, Some(bnc::DURATION_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Automatic {
                identifier: Some(identifier),
                secondary,
            })
        },
        (bnc::EXPLICIT_DECIMAL_FORMAT, Some(bnc::DECIMAL_CELL_FORMAT_KIND), Some(identifier))
            if secondary.is_none() =>
        {
            Ok(CellFormatReference::Explicit {
                identifier,
                secondary: None,
            })
        },
        (bnc::EXPLICIT_CURRENCY_FORMAT, Some(bnc::CURRENCY_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Explicit {
                identifier,
                secondary,
            })
        },
        (
            bnc::EXPLICIT_DATE_TIME_FORMAT,
            Some(bnc::DATE_TIME_CELL_FORMAT_KIND),
            Some(identifier),
        ) if secondary.is_none() => Ok(CellFormatReference::Explicit {
            identifier,
            secondary: None,
        }),
        (bnc::EXPLICIT_DURATION_FORMAT, Some(bnc::DURATION_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Explicit {
                identifier,
                secondary,
            })
        },
        (bnc::EXPLICIT_CHECKBOX_FORMAT, Some(bnc::CHECKBOX_CELL_FORMAT_KIND), Some(identifier))
            if secondary.is_none() =>
        {
            Ok(CellFormatReference::Explicit {
                identifier,
                secondary: None,
            })
        },
        _ => Err(Error::InvalidFormat(
            "Table cell contains inconsistent data-format metadata".to_owned(),
        )),
    }
}

fn resolve_format_table(
    package: &IWorkPackage,
    location: &model::CellLocation,
) -> Result<storage::ResolvedTableDataList> {
    let identifier = location
        .descriptor
        .model
        .base_data_store
        .format_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers table has no current format table".to_owned())
        })?;
    storage::resolve_table_data_list(
        package,
        &location.object_locations,
        identifier,
        tst::table_data_list::ListType::Format,
    )
}

fn required_format_entry(
    resolved: &storage::ResolvedTableDataList,
    identifier: u32,
) -> Result<&storage::LocatedTableDataListEntry> {
    let entry = resolved
        .entries
        .iter()
        .find(|located| located.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers format table has no entry for identifier {identifier}"
            ))
        })?;
    if entry.entry.refcount == 0 {
        return Err(Error::InvalidFormat(format!(
            "Numbers format-table entry {identifier} has a zero reference count"
        )));
    }
    if entry.entry.format.is_none() {
        return Err(Error::InvalidFormat(format!(
            "Numbers format-table entry {identifier} has no format payload"
        )));
    }
    Ok(entry)
}

fn ensure_current_format_table(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    if location
        .descriptor
        .model
        .base_data_store
        .format_table
        .is_some()
    {
        resolve_format_table(package, &location)?;
        return Ok(());
    }

    let identifier = location
        .descriptor
        .model
        .base_data_store
        .format_table_pre_bnc
        .identifier;
    storage::resolve_table_data_list(
        package,
        &location.object_locations,
        identifier,
        tst::table_data_list::ListType::Format,
    )?;
    let archive_name = location
        .object_locations
        .get(&location.descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table object {} is missing",
                location.descriptor.object_id
            ))
        })?
        .clone();
    package.update_archive(&archive_name, |archive| {
        let object = archive
            .object_mut(location.descriptor.object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table object {} is missing",
                    location.descriptor.object_id
                ))
            })?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6000 || message.type_ == 6001)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers table-model payload",
                    location.descriptor.object_id
                ))
            })?;
        let previous = TableModelArchive::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.base_data_store.format_table = Some(tsp::Reference {
            identifier,
            ..Default::default()
        });
        let data = storage::rewrite_table_model_format_table_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        if !object.archive_info.message_infos[index]
            .object_references
            .contains(&identifier)
        {
            object.archive_info.message_infos[index]
                .object_references
                .push(identifier);
        }
        Ok(())
    })
}

fn append_format_entry(
    package: &mut IWorkPackage,
    resolved: &storage::ResolvedTableDataList,
    format: FormatStructArchive,
) -> Result<u32> {
    let key = storage::next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(resolved.table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers format table object {} is missing",
                resolved.table_id
            ))
        })?;
        let index =
            storage::table_data_list_message_index(object, tst::table_data_list::ListType::Format)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {} has no Numbers format-list payload",
                        resolved.table_id
                    ))
                })?;
        let previous = TableDataList::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers format identifier overflow".to_owned()))?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            format: Some(format),
            ..Default::default()
        });
        let data = storage::rewrite_table_data_list_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })?;
    Ok(key)
}

#[cfg(test)]
fn number_format_to_native(format: TableCellNumberFormat) -> FormatStructArchive {
    data_format_to_native(&format.into()).expect("number format is explicit")
}

#[cfg(test)]
fn number_format_from_native(native: &FormatStructArchive) -> Result<TableCellNumberFormat> {
    match data_format_from_native(native)? {
        TableCellDataFormat::Number(format) => Ok(format),
        format => Err(Error::InvalidFormat(format!(
            "Expected a Number cell format, found {format:?}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::{CellValue, NumbersDocument, NumbersDocumentBuilder, NumbersEditor};
    use crate::table_cell_data_format::{
        TableCellDurationStyle, TableCellDurationUnit, TableCellDurationUnitRange,
        TableCellFractionAccuracy, TableCellFractionFormat, TableCellNumeralSystemBase,
        TableCellNumeralSystemFixedPlaces, TableCellNumeralSystemFormat,
        TableCellNumeralSystemNegativeStyle, TableCellNumeralSystemPlaces,
    };

    #[test]
    fn number_format_native_codec_is_strict_and_roundtrips() {
        let format = TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::Parentheses,
            TableCellThousandsSeparator::Shown,
        );
        assert_eq!(
            number_format_from_native(&number_format_to_native(format)).unwrap(),
            format
        );

        let mut invalid = number_format_to_native(format);
        invalid.decimal_places = Some(31);
        assert!(number_format_from_native(&invalid).is_err());
        invalid = number_format_to_native(format);
        invalid.format_type = Some(999);
        assert!(number_format_from_native(&invalid).is_err());

        let percentage = TableCellPercentageFormat::new(
            TableCellDecimalPlaces::fixed(3).unwrap(),
            TableCellNegativeNumberStyle::RedParentheses,
            TableCellThousandsSeparator::Hidden,
        );
        let native = data_format_to_native(&percentage.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_PERCENTAGE_FORMAT_TYPE));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Percentage(percentage)
        );

        let currency = TableCellCurrencyFormat::new(
            TableCellCurrencyCode::EUR,
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::Parentheses,
            TableCellThousandsSeparator::Shown,
            TableCellCurrencyStyle::Accounting,
        );
        let mut native = data_format_to_native(&currency.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_CURRENCY_FORMAT_TYPE));
        assert_eq!(native.currency_code.as_deref(), Some("EUR"));
        assert_eq!(native.use_accounting_style, Some(true));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Currency(currency)
        );

        native.currency_code = Some("euro".to_owned());
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&currency.into()).unwrap();
        native.use_accounting_style = None;
        assert!(data_format_from_native(&native).is_err());

        let scientific =
            TableCellScientificFormat::new(TableCellFixedDecimalPlaces::new(5).unwrap());
        let mut native = data_format_to_native(&scientific.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_SCIENTIFIC_FORMAT_TYPE));
        assert_eq!(native.decimal_places, Some(5));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Scientific(scientific)
        );
        native.negative_style = Some(2);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&scientific.into()).unwrap();
        native.decimal_places = Some(NATIVE_AUTOMATIC_DECIMAL_PLACES);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&scientific.into()).unwrap();
        native.fraction_accuracy = Some(8);
        assert!(data_format_from_native(&native).is_err());

        let accuracies = [
            (TableCellFractionAccuracy::UpToOneDigit, u32::MAX),
            (TableCellFractionAccuracy::UpToTwoDigits, u32::MAX - 1),
            (TableCellFractionAccuracy::UpToThreeDigits, u32::MAX - 2),
            (TableCellFractionAccuracy::Halves, 2),
            (TableCellFractionAccuracy::Quarters, 4),
            (TableCellFractionAccuracy::Eighths, 8),
            (TableCellFractionAccuracy::Sixteenths, 16),
            (TableCellFractionAccuracy::Tenths, 10),
            (TableCellFractionAccuracy::Hundredths, 100),
        ];
        for (accuracy, native_accuracy) in accuracies {
            let fraction = TableCellFractionFormat::new(accuracy);
            let native = data_format_to_native(&fraction.into()).unwrap();
            assert_eq!(native.format_type, Some(NATIVE_FRACTION_FORMAT_TYPE));
            assert_eq!(native.fraction_accuracy, Some(native_accuracy));
            assert_eq!(
                data_format_from_native(&native).unwrap(),
                TableCellDataFormat::Fraction(fraction)
            );
        }
        let mut invalid =
            data_format_to_native(&TableCellFractionFormat::default().into()).unwrap();
        invalid.fraction_accuracy = Some(3);
        assert!(data_format_from_native(&invalid).is_err());
        invalid = data_format_to_native(&TableCellFractionFormat::default().into()).unwrap();
        invalid.decimal_places = Some(2);
        assert!(data_format_from_native(&invalid).is_err());

        let numeral_system = TableCellNumeralSystemFormat::new(
            TableCellNumeralSystemBase::HEXADECIMAL,
            TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
            TableCellNumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        let mut native = data_format_to_native(&numeral_system.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE));
        assert_eq!(native.base, Some(16));
        assert_eq!(native.base_places, Some(8));
        assert_eq!(native.base_use_minus_sign, Some(false));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::NumeralSystem(numeral_system)
        );
        native.base = Some(37);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&numeral_system.into()).unwrap();
        native.base_places = Some(33);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&numeral_system.into()).unwrap();
        native.base = Some(10);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&numeral_system.into()).unwrap();
        native.decimal_places = Some(2);
        assert!(data_format_from_native(&native).is_err());

        let date_time = TableCellDateTimeFormat::iso_date_time_24_hour_with_seconds();
        let mut native =
            data_format_to_native(&TableCellDataFormat::DateTime(date_time.clone())).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_DATE_TIME_FORMAT_TYPE));
        assert_eq!(
            native.date_time_format.as_deref(),
            Some("yyyy-MM-dd H:mm:ss")
        );
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::DateTime(date_time.clone())
        );
        native.date_time_format = None;
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&TableCellDataFormat::DateTime(date_time.clone())).unwrap();
        native.suppress_time_format = Some(false);
        assert!(data_format_from_native(&native).is_err());

        let range = TableCellDurationUnitRange::new(
            TableCellDurationUnit::Hours,
            TableCellDurationUnit::Milliseconds,
        )
        .unwrap();
        let duration = TableCellDurationFormat::custom(TableCellDurationStyle::Abbreviated, range);
        let mut native = data_format_to_native(&duration.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_DURATION_FORMAT_TYPE));
        assert_eq!(native.duration_style, Some(1));
        assert_eq!(native.duration_unit_largest, Some(4));
        assert_eq!(native.duration_unit_smallest, Some(32));
        assert_eq!(native.use_automatic_duration_units, Some(false));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Duration(duration)
        );
        native.duration_style = Some(3);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&duration.into()).unwrap();
        native.duration_unit_largest = Some(3);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&duration.into()).unwrap();
        native.use_automatic_duration_units = None;
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&duration.into()).unwrap();
        native.decimal_places = Some(2);
        assert!(data_format_from_native(&native).is_err());

        let checkbox = TableCellDataFormat::Checkbox(TableCellCheckboxFormat);
        let mut native = data_format_to_native(&checkbox).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_CHECKBOX_FORMAT_TYPE));
        assert_eq!(data_format_from_native(&native).unwrap(), checkbox);
        native.bool_true_string = Some("Yes".to_owned());
        assert!(data_format_from_native(&native).is_err());

        let star_rating = TableCellDataFormat::StarRating(TableCellStarRatingFormat);
        let mut native = data_format_to_native(&star_rating).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_STAR_RATING_FORMAT_TYPE));
        assert_eq!(data_format_from_native(&native).unwrap(), star_rating);
        native.control_maximum = Some(5.0);
        assert!(data_format_from_native(&native).is_err());

        let numeric_control_displays = [
            TableCellNumericControlDisplayFormat::Number(TableCellNumberFormat::default()),
            TableCellNumericControlDisplayFormat::Currency(TableCellCurrencyFormat::default()),
            TableCellNumericControlDisplayFormat::Percentage(TableCellPercentageFormat::default()),
            TableCellNumericControlDisplayFormat::Fraction(TableCellFractionFormat::default()),
            TableCellNumericControlDisplayFormat::Scientific(TableCellScientificFormat::default()),
            TableCellNumericControlDisplayFormat::NumeralSystem(
                TableCellNumeralSystemFormat::default(),
            ),
        ];
        for display in numeric_control_displays {
            let native = numeric_control_display_to_native(&display).unwrap();
            assert_eq!(
                numeric_control_display_from_native(&native).unwrap(),
                display
            );
        }
        let invalid_numeric_control_native =
            data_format_to_native(&TableCellDataFormat::DateTime(date_time)).unwrap();
        assert!(numeric_control_display_from_native(&invalid_numeric_control_native).is_err());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_checkbox_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Checkboxes")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Boolean(true))
            .unwrap();
        editor
            .set_table_cell_checkbox_format(table_id, 1, 1, TableCellCheckboxFormat)
            .unwrap();
        editor
            .set_table_cell_checkbox_format(table_id, 1, 2, TableCellCheckboxFormat)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        let control_table_id = location
            .descriptor
            .model
            .base_data_store
            .control_cell_spec_table
            .as_ref()
            .unwrap()
            .identifier;
        let controls = storage::resolve_table_data_list(
            editor.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert_eq!(controls.entries.len(), 1);
        assert_eq!(controls.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_checkbox_format(table_id, 1, 1).unwrap(),
            Some(TableCellCheckboxFormat)
        );
        let document = NumbersDocument::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 2),
            Some(&CellValue::Boolean(false))
        );
        assert!(
            reopened
                .reset_table_cell_checkbox_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_checkbox_format(table_id, 1, 2).unwrap(),
            Some(TableCellCheckboxFormat)
        );
        assert!(
            reopened
                .reset_table_cell_checkbox_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert!(controls.entries.is_empty());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_star_rating_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Ratings")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(3.0))
            .unwrap();
        editor
            .set_table_cell_star_rating_format(table_id, 1, 1, TableCellStarRatingFormat)
            .unwrap();
        editor
            .set_table_cell_star_rating_format(table_id, 1, 2, TableCellStarRatingFormat)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        let control_table_id = location
            .descriptor
            .model
            .base_data_store
            .control_cell_spec_table
            .as_ref()
            .unwrap()
            .identifier;
        let controls = storage::resolve_table_data_list(
            editor.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert_eq!(controls.entries.len(), 1);
        assert_eq!(controls.entries[0].entry.refcount, 2);
        assert_eq!(
            controls.entries[0].entry.cell_spec,
            Some(tst::CellSpecArchive {
                interaction_type: 6,
                range_control_min: Some(0.0),
                range_control_max: Some(5.0),
                range_control_inc: Some(1.0),
                ..Default::default()
            })
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_star_rating_format(table_id, 1, 1)
                .unwrap(),
            Some(TableCellStarRatingFormat)
        );
        let document = NumbersDocument::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 2),
            Some(&CellValue::Number(0.0))
        );
        assert!(
            reopened
                .reset_table_cell_star_rating_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened
                .table_cell_star_rating_format(table_id, 1, 2)
                .unwrap(),
            Some(TableCellStarRatingFormat)
        );
        assert!(
            reopened
                .reset_table_cell_star_rating_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert!(controls.entries.is_empty());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_slider_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Sliders")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let range = TableCellSliderRange::new(-10.0, 30.0, 0.5).unwrap();
        let number_display = TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
        );
        let number_slider = TableCellSliderFormat::new(range, number_display.into());
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(25.0))
            .unwrap();
        editor
            .set_table_cell_slider_format(table_id, 1, 1, number_slider.clone())
            .unwrap();
        editor
            .set_table_cell_slider_format(table_id, 1, 2, number_slider.clone())
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        let control_table_id = location
            .descriptor
            .model
            .base_data_store
            .control_cell_spec_table
            .as_ref()
            .unwrap()
            .identifier;
        let controls = storage::resolve_table_data_list(
            editor.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert_eq!(controls.entries.len(), 1);
        assert_eq!(controls.entries[0].entry.refcount, 2);
        assert_eq!(
            controls.entries[0].entry.cell_spec,
            Some(tst::CellSpecArchive {
                interaction_type: 5,
                range_control_min: Some(-10.0),
                range_control_max: Some(30.0),
                range_control_inc: Some(0.5),
                ..Default::default()
            })
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_slider_format(table_id, 1, 1).unwrap(),
            Some(number_slider.clone())
        );
        let document = NumbersDocument::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 2),
            Some(&CellValue::Number(10.0))
        );

        let currency_slider =
            TableCellSliderFormat::new(range, TableCellCurrencyFormat::default().into());
        reopened
            .set_table_cell_slider_format(table_id, 1, 1, currency_slider.clone())
            .unwrap();
        assert_eq!(
            reopened.table_cell_slider_format(table_id, 1, 1).unwrap(),
            Some(currency_slider)
        );
        assert!(
            reopened
                .reset_table_cell_slider_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_slider_format(table_id, 1, 2).unwrap(),
            Some(number_slider)
        );
        assert!(
            reopened
                .reset_table_cell_slider_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert!(controls.entries.is_empty());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_stepper_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Steppers")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let range = TableCellStepperRange::new(-10.0, 30.0, 0.5).unwrap();
        let number_display = TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
        );
        let number_stepper = TableCellStepperFormat::new(range, number_display.into());
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(25.0))
            .unwrap();
        editor
            .set_table_cell_stepper_format(table_id, 1, 1, number_stepper.clone())
            .unwrap();
        editor
            .set_table_cell_stepper_format(table_id, 1, 2, number_stepper.clone())
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        let control_table_id = location
            .descriptor
            .model
            .base_data_store
            .control_cell_spec_table
            .as_ref()
            .unwrap()
            .identifier;
        let controls = storage::resolve_table_data_list(
            editor.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert_eq!(controls.entries.len(), 1);
        assert_eq!(controls.entries[0].entry.refcount, 2);
        assert_eq!(
            controls.entries[0].entry.cell_spec,
            Some(tst::CellSpecArchive {
                interaction_type: 4,
                range_control_min: Some(-10.0),
                range_control_max: Some(30.0),
                range_control_inc: Some(0.5),
                ..Default::default()
            })
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_stepper_format(table_id, 1, 1).unwrap(),
            Some(number_stepper.clone())
        );
        let document = NumbersDocument::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 2),
            Some(&CellValue::Number(-10.0))
        );

        let currency_stepper =
            TableCellStepperFormat::new(range, TableCellCurrencyFormat::default().into());
        reopened
            .set_table_cell_stepper_format(table_id, 1, 1, currency_stepper.clone())
            .unwrap();
        assert_eq!(
            reopened.table_cell_stepper_format(table_id, 1, 1).unwrap(),
            Some(currency_stepper)
        );
        assert!(
            reopened
                .reset_table_cell_stepper_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_stepper_format(table_id, 1, 2).unwrap(),
            Some(number_stepper)
        );
        assert!(
            reopened
                .reset_table_cell_stepper_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert!(controls.entries.is_empty());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_duration_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Durations")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let range = TableCellDurationUnitRange::hours_to_milliseconds();
        let duration = TableCellDurationFormat::custom(TableCellDurationStyle::Abbreviated, range);
        editor
            .set_cell(table_id, 1, 1, CellValue::Duration(3_723.5))
            .unwrap();
        editor
            .set_cell(table_id, 1, 2, CellValue::Number(1.5))
            .unwrap();
        editor
            .set_table_cell_duration_format(table_id, 1, 1, duration)
            .unwrap();
        editor
            .set_table_cell_duration_format(table_id, 1, 2, duration)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_duration_format(table_id, 1, 1).unwrap(),
            Some(duration)
        );
        let document = NumbersDocument::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 2),
            Some(&CellValue::Duration(129_600.0))
        );
        assert!(
            reopened
                .reset_table_cell_duration_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_duration_format(table_id, 1, 2).unwrap(),
            Some(duration)
        );
        assert!(
            reopened
                .reset_table_cell_duration_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn native_automatic_duration_reference_is_recognized() {
        let native = [
            0x05, 0x07, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x10, 0x01, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x17, 0xad, 0x40, 0x04, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00,
        ];
        let cell = BncCell::parse(&native).unwrap();
        let CellFormatReference::Automatic {
            identifier,
            secondary,
        } = format_reference(&cell).unwrap()
        else {
            panic!("native automatic Duration cell was treated as explicit");
        };
        assert_eq!(identifier, Some(8));
        assert_eq!(secondary, None);
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_date_time_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Dates")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let date_time = TableCellDateTimeFormat::iso_date_time_24_hour_with_seconds();
        editor
            .set_cell(table_id, 1, 1, CellValue::Date(789_332_889.0))
            .unwrap();
        editor
            .set_cell(table_id, 1, 2, CellValue::Date(789_332_889.0))
            .unwrap();
        editor
            .set_table_cell_date_time_format(table_id, 1, 1, date_time.clone())
            .unwrap();
        editor
            .set_table_cell_date_time_format(table_id, 1, 2, date_time.clone())
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_date_time_format(table_id, 1, 1)
                .unwrap(),
            Some(date_time.clone())
        );
        assert!(
            reopened
                .reset_table_cell_date_time_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened
                .table_cell_date_time_format(table_id, 1, 2)
                .unwrap(),
            Some(date_time)
        );
        assert!(
            reopened
                .reset_table_cell_date_time_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_numeral_system_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Numeral Systems")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let numeral_system = TableCellNumeralSystemFormat::new(
            TableCellNumeralSystemBase::HEXADECIMAL,
            TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
            TableCellNumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_numeral_system_format(table_id, 1, 1, numeral_system)
            .unwrap();
        editor
            .set_table_cell_numeral_system_format(table_id, 1, 2, numeral_system)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_numeral_system_format(table_id, 1, 1)
                .unwrap(),
            Some(numeral_system)
        );
        assert!(
            reopened
                .reset_table_cell_numeral_system_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened
                .table_cell_numeral_system_format(table_id, 1, 2)
                .unwrap(),
            Some(numeral_system)
        );
        assert!(
            reopened
                .reset_table_cell_numeral_system_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_fraction_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Fractions")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let fraction = TableCellFractionFormat::new(TableCellFractionAccuracy::Eighths);
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(-12.375))
            .unwrap();
        editor
            .set_table_cell_fraction_format(table_id, 1, 1, fraction)
            .unwrap();
        editor
            .set_table_cell_fraction_format(table_id, 1, 2, fraction)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_fraction_format(table_id, 1, 1).unwrap(),
            Some(fraction)
        );
        assert!(
            reopened
                .reset_table_cell_fraction_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_fraction_format(table_id, 1, 2).unwrap(),
            Some(fraction)
        );
        assert!(
            reopened
                .reset_table_cell_fraction_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_scientific_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Scientific")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let scientific =
            TableCellScientificFormat::new(TableCellFixedDecimalPlaces::new(5).unwrap());
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_scientific_format(table_id, 1, 1, scientific)
            .unwrap();
        editor
            .set_table_cell_scientific_format(table_id, 1, 2, scientific)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_scientific_format(table_id, 1, 1)
                .unwrap(),
            Some(scientific)
        );
        assert!(
            reopened
                .reset_table_cell_scientific_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened
                .table_cell_scientific_format(table_id, 1, 2)
                .unwrap(),
            Some(scientific)
        );
        assert!(
            reopened
                .reset_table_cell_scientific_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_resets_currency_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Currencies")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let currency = TableCellCurrencyFormat::new(
            TableCellCurrencyCode::USD,
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::Parentheses,
            TableCellThousandsSeparator::Shown,
            TableCellCurrencyStyle::Accounting,
        );
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(-1_234.5))
            .unwrap();
        editor
            .set_table_cell_currency_format(table_id, 1, 1, currency)
            .unwrap();
        editor
            .set_table_cell_currency_format(table_id, 1, 2, currency)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_currency_format(table_id, 1, 1).unwrap(),
            Some(currency)
        );
        assert!(
            reopened
                .reset_table_cell_currency_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_currency_format(table_id, 1, 2).unwrap(),
            Some(currency)
        );
        assert!(
            reopened
                .reset_table_cell_currency_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn source_built_table_roundtrips_and_reuses_number_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Formats")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let format = TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::Parentheses,
            TableCellThousandsSeparator::Shown,
        );
        assert_eq!(
            editor.table_cell_number_format(table_id, 1, 1).unwrap(),
            None
        );
        editor
            .set_cell(table_id, 1, 1, CellValue::Number(1_234.5))
            .unwrap();
        editor
            .set_table_cell_number_format(table_id, 1, 1, format)
            .unwrap();
        editor
            .set_table_cell_number_format(table_id, 1, 2, format)
            .unwrap();
        editor
            .set_table_cell_number_format(table_id, 1, 2, format)
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_number_format(table_id, 1, 1).unwrap(),
            Some(format)
        );
        assert!(
            reopened
                .reset_table_cell_number_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_number_format(table_id, 1, 2).unwrap(),
            Some(format)
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let formats = resolve_format_table(reopened.package(), &location).unwrap();
        assert_eq!(formats.entries[0].entry.refcount, 1);

        assert!(
            reopened
                .reset_table_cell_number_format(table_id, 1, 2)
                .unwrap()
        );
        assert!(
            !reopened
                .reset_table_cell_number_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
    }

    #[test]
    fn invalid_number_format_coordinate_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell_number_format(table_id, 2, 0, TableCellNumberFormat::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn source_built_table_converts_and_reuses_percentage_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Percentages")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let number = TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(1).unwrap(),
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
        );
        let percentage = TableCellPercentageFormat::new(
            TableCellDecimalPlaces::fixed(2).unwrap(),
            TableCellNegativeNumberStyle::Parentheses,
            TableCellThousandsSeparator::Shown,
        );
        editor
            .set_table_cell_data_format(table_id, 1, 1, number.into())
            .unwrap();
        editor
            .set_cell(table_id, 1, 2, CellValue::Number(-12.345))
            .unwrap();
        assert_eq!(
            editor.table_cell_data_format(table_id, 1, 2).unwrap(),
            TableCellDataFormat::Automatic
        );
        editor
            .set_table_cell_percentage_format(table_id, 1, 1, percentage)
            .unwrap();
        editor
            .set_table_cell_percentage_format(table_id, 1, 2, percentage)
            .unwrap();

        assert_eq!(
            editor.table_cell_data_format(table_id, 1, 1).unwrap(),
            TableCellDataFormat::Percentage(percentage)
        );
        assert!(editor.table_cell_number_format(table_id, 1, 1).is_err());
        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        assert_eq!(
            formats.entries[0]
                .entry
                .format
                .as_ref()
                .and_then(|format| format.format_type),
            Some(NATIVE_PERCENTAGE_FORMAT_TYPE)
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_percentage_format(table_id, 1, 1)
                .unwrap(),
            Some(percentage)
        );
        reopened
            .set_table_cell_data_format(table_id, 1, 1, TableCellDataFormat::Automatic)
            .unwrap();
        assert_eq!(
            reopened.table_cell_data_format(table_id, 1, 1).unwrap(),
            TableCellDataFormat::Automatic
        );
        assert_eq!(
            reopened
                .table_cell_percentage_format(table_id, 1, 2)
                .unwrap(),
            Some(percentage)
        );
    }
}
