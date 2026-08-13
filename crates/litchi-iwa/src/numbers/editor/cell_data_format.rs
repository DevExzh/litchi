//! Typed data-format storage for native table cells.

use super::*;
use crate::numbers::bnc;
use crate::protobuf::tsk::FormatStructArchive;
use litchi_numbers::cell::data_format::control::DisplayFormat;
use litchi_numbers::cell::data_format::pop_up_menu::InitialSelection;
use litchi_numbers::cell::data_format::{
    Checkbox, Currency, Custom, DataFormat, DateTime, Duration, Fraction, Number, NumeralSystem,
    Percentage, PopUpMenu, Scientific, Slider, StarRating, Stepper, Text,
};

macro_rules! map_semantic_format_error {
    ($error:ty) => {
        impl From<$error> for crate::Error {
            fn from(error: $error) -> Self {
                Self::InvalidFormat(error.to_string())
            }
        }
    };
}

map_semantic_format_error!(litchi_numbers::cell::data_format::control::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::custom::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::date_time::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::duration::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::number::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::numeral_system::Error);
map_semantic_format_error!(litchi_numbers::cell::data_format::pop_up_menu::Error);

mod codec;
mod control;
mod custom;
mod pop_up_menu;
use codec::*;

pub(super) fn cell_data_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<DataFormat> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(DataFormat::Automatic);
    };
    let cell = BncCell::parse(&data)?;
    match format_reference(&cell)? {
        CellFormatReference::Automatic { .. } => Ok(DataFormat::Automatic),
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
                    let format = if custom::reference_uuid(native)?.is_some() {
                        DataFormat::Custom(custom::resolve_reference(package, native)?)
                    } else {
                        data_format_from_native(native)?
                    };
                    if matches!(
                        format,
                        DataFormat::Checkbox(_)
                            | DataFormat::StarRating(_)
                            | DataFormat::Slider(_)
                            | DataFormat::Stepper(_)
                            | DataFormat::PopUpMenu(_)
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
                    if !matches!(format, DataFormat::Checkbox(_)) {
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
                    if !matches!(format, DataFormat::StarRating(_)) {
                        return Err(Error::InvalidFormat(
                            "Star Rating control references a non-Star-Rating format".to_owned(),
                        ));
                    }
                    Ok(format)
                },
                Some(control::ControlCellSpecKind::Slider(range)) => {
                    let display_format = numeric_control_display_from_native(native)?;
                    let expected_kind = match display_format {
                        DisplayFormat::Currency(_) => bnc::CURRENCY_CELL_FORMAT_KIND,
                        _ => bnc::DECIMAL_CELL_FORMAT_KIND,
                    };
                    if cell.cell_format_kind() != Some(expected_kind) {
                        return Err(Error::InvalidFormat(
                            "Slider control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    Ok(DataFormat::Slider(Slider::new(range, display_format)))
                },
                Some(control::ControlCellSpecKind::Stepper(range)) => {
                    let display_format = numeric_control_display_from_native(native)?;
                    let expected_kind = match display_format {
                        DisplayFormat::Currency(_) => bnc::CURRENCY_CELL_FORMAT_KIND,
                        _ => bnc::DECIMAL_CELL_FORMAT_KIND,
                    };
                    if cell.cell_format_kind() != Some(expected_kind) {
                        return Err(Error::InvalidFormat(
                            "Stepper control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    Ok(DataFormat::Stepper(Stepper::new(range, display_format)))
                },
                Some(control::ControlCellSpecKind::PopUpMenu(format)) => {
                    validate_text_format(native)?;
                    if cell.cell_format_kind() != Some(bnc::TEXT_CELL_FORMAT_KIND) {
                        return Err(Error::InvalidFormat(
                            "Pop-Up Menu control uses inconsistent BNC format metadata".to_owned(),
                        ));
                    }
                    Ok(DataFormat::PopUpMenu(format))
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
) -> Result<Option<Number>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Number(format) => Ok(Some(format)),
        DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Number data format".to_owned(),
        )),
    }
}

pub(super) fn cell_custom_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Custom>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Custom(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use a Custom data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_custom_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(false),
        DataFormat::Custom(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Custom format from a non-Custom cell".to_owned(),
        )),
    }
}

pub(super) fn cell_text_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Text>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Text(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Text data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_text_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(false),
        DataFormat::Text(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Text format from a non-Text cell".to_owned(),
        )),
    }
}

pub(super) fn cell_currency_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Currency>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Currency(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Currency(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Currency format from a non-Currency cell".to_owned(),
        )),
    }
}

pub(super) fn cell_percentage_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Percentage>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Percentage(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Table cell does not use the Percentage data format".to_owned(),
        )),
    }
}

pub(super) fn cell_scientific_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Scientific>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Scientific(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Scientific(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Scientific format from a non-Scientific cell".to_owned(),
        )),
    }
}

pub(super) fn cell_fraction_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Fraction>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Fraction(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Fraction(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Fraction format from a non-Fraction cell".to_owned(),
        )),
    }
}

pub(super) fn cell_numeral_system_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<NumeralSystem>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::NumeralSystem(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::NumeralSystem(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Numeral System format from a non-Numeral-System cell".to_owned(),
        )),
    }
}

pub(super) fn cell_date_time_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<DateTime>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::DateTime(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::DateTime(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Date & Time format from a non-Date-Time cell".to_owned(),
        )),
    }
}

pub(super) fn cell_duration_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Duration>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Duration(format) => Ok(Some(format)),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Duration(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
            "Cannot reset Duration format from a non-Duration cell".to_owned(),
        )),
    }
}

pub(super) fn cell_checkbox_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Checkbox>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Checkbox(format) => Ok(Some(format)),
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Checkbox(_) => reset_cell_data_format(package, table_id, row, column),
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
) -> Result<Option<StarRating>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::StarRating(format) => Ok(Some(format)),
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
        DataFormat::Automatic => Ok(false),
        DataFormat::StarRating(_) => reset_cell_data_format(package, table_id, row, column),
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
) -> Result<Option<Slider>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Slider(format) => Ok(Some(format)),
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Slider(_) => reset_cell_data_format(package, table_id, row, column),
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
) -> Result<Option<Stepper>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::Stepper(format) => Ok(Some(format)),
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Stepper(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Stepper format from a non-Stepper cell".to_owned(),
        )),
    }
}

pub(super) fn cell_pop_up_menu_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<PopUpMenu>> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(None),
        DataFormat::PopUpMenu(format) => Ok(Some(format)),
        _ => Err(Error::InvalidFormat(
            "Table cell does not use the Pop-Up Menu data format".to_owned(),
        )),
    }
}

pub(super) fn reset_cell_pop_up_menu_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    match cell_data_format(package, table_id, row, column)? {
        DataFormat::Automatic => Ok(false),
        DataFormat::PopUpMenu(_) => reset_cell_data_format(package, table_id, row, column),
        _ => Err(Error::InvalidFormat(
            "Cannot reset Pop-Up Menu format from a non-Pop-Up-Menu cell".to_owned(),
        )),
    }
}

pub(super) fn set_cell_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: Number,
) -> Result<()> {
    set_cell_data_format(package, table_id, row, column, &format.into())
}

pub(super) fn set_cell_data_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: &DataFormat,
) -> Result<()> {
    if format == &DataFormat::Automatic {
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
        DataFormat::Currency(_) => bnc::CellDataFormatKind::Currency,
        DataFormat::DateTime(_) => bnc::CellDataFormatKind::DateTime,
        DataFormat::Duration(_) => bnc::CellDataFormatKind::Duration,
        DataFormat::Text(_) => bnc::CellDataFormatKind::Text,
        DataFormat::Checkbox(_) => bnc::CellDataFormatKind::Checkbox,
        DataFormat::StarRating(_) => bnc::CellDataFormatKind::StarRating,
        DataFormat::Slider(format) => match format.display_format() {
            DisplayFormat::Currency(_) => bnc::CellDataFormatKind::NumericControlCurrency,
            _ => bnc::CellDataFormatKind::NumericControlNumberOrPercentage,
        },
        DataFormat::Stepper(format) => match format.display_format() {
            DisplayFormat::Currency(_) => bnc::CellDataFormatKind::NumericControlCurrency,
            _ => bnc::CellDataFormatKind::NumericControlNumberOrPercentage,
        },
        DataFormat::PopUpMenu(_) => bnc::CellDataFormatKind::PopUpMenu,
        DataFormat::Custom(format) => custom::scalar_kind(format),
        DataFormat::Number(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_) => bnc::CellDataFormatKind::NumberOrPercentage,
        DataFormat::Automatic => unreachable!("handled above"),
    };
    let native = match format {
        DataFormat::Custom(format) => custom::acquire_reference(package, format)?,
        _ => data_format_to_native(format)?,
    };
    let new_control_identifier = match format {
        DataFormat::Checkbox(_) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Checkbox,
        )?),
        DataFormat::StarRating(_) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::StarRating,
        )?),
        DataFormat::Slider(format) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Slider(format.range()),
        )?),
        DataFormat::Stepper(format) => Some(control::acquire_spec(
            package,
            &location,
            old_control_identifier,
            control::ControlCellSpecKind::Stepper(format.range()),
        )?),
        DataFormat::PopUpMenu(format) => Some(control::acquire_pop_up_menu_spec(
            package,
            &location,
            old_control_identifier,
            format,
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
    let mut old_custom_references = Vec::new();
    for entry in &old_entries {
        if let Some(native) = entry.entry.format.as_ref()
            && custom::reference_uuid(native)?.is_some()
        {
            old_custom_references.push((entry.entry.key, native.clone()));
        }
    }

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

    if let DataFormat::Slider(format) = format
        && cell.cached_scalar()?.is_none()
    {
        cell.set_plain_number(format.range().midpoint())?;
    }
    if let DataFormat::Stepper(format) = format
        && cell.cached_scalar()?.is_none()
    {
        cell.set_plain_number(format.range().minimum())?;
    }
    if let DataFormat::PopUpMenu(format) = format
        && matches!(format.initial_selection(), InitialSelection::FirstItem)
        && matches!(cell.stored_value(), StoredValue::Empty)
    {
        let identifier = storage::update_string_table(
            package,
            &location.object_locations,
            location
                .descriptor
                .model
                .base_data_store
                .string_table
                .identifier,
            None,
            Some(format.items()[0].as_str()),
        )?
        .expect("new text produces a string-table identifier");
        cell.set_string(identifier);
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
    for (identifier, reference) in old_custom_references {
        if identifier != new_identifier {
            custom::release_reference_if_unused(package, &reference)?;
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Number(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Currency(_)
        | DataFormat::Percentage(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
        DataFormat::Automatic => Ok(false),
        DataFormat::Percentage(_) => reset_cell_data_format(package, table_id, row, column),
        DataFormat::Number(_)
        | DataFormat::Currency(_)
        | DataFormat::Scientific(_)
        | DataFormat::Fraction(_)
        | DataFormat::NumeralSystem(_)
        | DataFormat::DateTime(_)
        | DataFormat::Duration(_)
        | DataFormat::Checkbox(_)
        | DataFormat::StarRating(_)
        | DataFormat::Slider(_)
        | DataFormat::Stepper(_)
        | DataFormat::Text(_)
        | DataFormat::Custom(_)
        | DataFormat::PopUpMenu(_) => Err(Error::InvalidFormat(
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
    let mut old_custom_references = Vec::new();
    for entry in &old_entries {
        if let Some(native) = entry.entry.format.as_ref()
            && custom::reference_uuid(native)?.is_some()
        {
            old_custom_references.push(native.clone());
        }
    }
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
    for reference in old_custom_references {
        custom::release_reference_if_unused(package, &reference)?;
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
        | (Some(bnc::TEXT_CELL_FORMAT_KIND), Some(_))
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
        (
            bnc::EXPLICIT_TEXT_FORMAT | bnc::EXPLICIT_CONVERTED_TEXT_FORMAT,
            Some(bnc::TEXT_CELL_FORMAT_KIND),
            Some(identifier),
        ) if secondary.is_none() => Ok(CellFormatReference::Explicit {
            identifier,
            secondary: None,
        }),
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
fn number_format_to_native(format: Number) -> FormatStructArchive {
    data_format_to_native(&format.into()).expect("number format is explicit")
}

#[cfg(test)]
fn number_format_from_native(native: &FormatStructArchive) -> Result<Number> {
    match data_format_from_native(native)? {
        DataFormat::Number(format) => Ok(format),
        format => Err(Error::InvalidFormat(format!(
            "Expected a Number cell format, found {format:?}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::cell::CellValue;
    use crate::numbers::editor::compatibility_document_from_bytes;
    use crate::numbers::{NumbersDocumentBuilder, NumbersEditor, SemanticTableCellAssertions};
    use litchi_numbers::cell::data_format::control::Range;
    use litchi_numbers::cell::data_format::custom::{
        Condition, ConditionValue, Custom, DateTime as CustomDateTime, DateTimePattern, Name,
        Number as CustomNumber, NumberPattern, NumberRule, Text as CustomText,
    };
    use litchi_numbers::cell::data_format::duration::{Style, Unit, UnitRange};
    use litchi_numbers::cell::data_format::number::{
        CurrencyCode, CurrencyStyle, DecimalPlaces, FixedDecimalPlaces, FractionAccuracy,
        NegativeStyle, ThousandsSeparator,
    };
    use litchi_numbers::cell::data_format::numeral_system::{
        Base, FixedPlaces, NegativeStyle as NumeralSystemNegativeStyle, Places,
    };

    #[test]
    fn number_format_native_codec_is_strict_and_roundtrips() {
        let format = Number::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
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

        let percentage = Percentage::new(
            DecimalPlaces::fixed(3).unwrap(),
            NegativeStyle::RedParentheses,
            ThousandsSeparator::Hidden,
        );
        let native = data_format_to_native(&percentage.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_PERCENTAGE_FORMAT_TYPE));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::Percentage(percentage)
        );

        let currency = Currency::new(
            CurrencyCode::EUR,
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
            CurrencyStyle::Accounting,
        );
        let mut native = data_format_to_native(&currency.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_CURRENCY_FORMAT_TYPE));
        assert_eq!(native.currency_code.as_deref(), Some("EUR"));
        assert_eq!(native.use_accounting_style, Some(true));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::Currency(currency)
        );

        native.currency_code = Some("euro".to_owned());
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&currency.into()).unwrap();
        native.use_accounting_style = None;
        assert!(data_format_from_native(&native).is_err());

        let scientific = Scientific::new(FixedDecimalPlaces::new(5).unwrap());
        let mut native = data_format_to_native(&scientific.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_SCIENTIFIC_FORMAT_TYPE));
        assert_eq!(native.decimal_places, Some(5));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::Scientific(scientific)
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
            (FractionAccuracy::UpToOneDigit, u32::MAX),
            (FractionAccuracy::UpToTwoDigits, u32::MAX - 1),
            (FractionAccuracy::UpToThreeDigits, u32::MAX - 2),
            (FractionAccuracy::Halves, 2),
            (FractionAccuracy::Quarters, 4),
            (FractionAccuracy::Eighths, 8),
            (FractionAccuracy::Sixteenths, 16),
            (FractionAccuracy::Tenths, 10),
            (FractionAccuracy::Hundredths, 100),
        ];
        for (accuracy, native_accuracy) in accuracies {
            let fraction = Fraction::new(accuracy);
            let native = data_format_to_native(&fraction.into()).unwrap();
            assert_eq!(native.format_type, Some(NATIVE_FRACTION_FORMAT_TYPE));
            assert_eq!(native.fraction_accuracy, Some(native_accuracy));
            assert_eq!(
                data_format_from_native(&native).unwrap(),
                DataFormat::Fraction(fraction)
            );
        }
        let mut invalid = data_format_to_native(&Fraction::default().into()).unwrap();
        invalid.fraction_accuracy = Some(3);
        assert!(data_format_from_native(&invalid).is_err());
        invalid = data_format_to_native(&Fraction::default().into()).unwrap();
        invalid.decimal_places = Some(2);
        assert!(data_format_from_native(&invalid).is_err());

        let numeral_system = NumeralSystem::new(
            Base::HEXADECIMAL,
            Places::Fixed(FixedPlaces::EIGHT),
            NumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        let mut native = data_format_to_native(&numeral_system.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_NUMERAL_SYSTEM_FORMAT_TYPE));
        assert_eq!(native.base, Some(16));
        assert_eq!(native.base_places, Some(8));
        assert_eq!(native.base_use_minus_sign, Some(false));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::NumeralSystem(numeral_system)
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

        let date_time = DateTime::iso_date_time_24_hour_with_seconds();
        let mut native = data_format_to_native(&DataFormat::DateTime(date_time.clone())).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_DATE_TIME_FORMAT_TYPE));
        assert_eq!(
            native.date_time_format.as_deref(),
            Some("yyyy-MM-dd H:mm:ss")
        );
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::DateTime(date_time.clone())
        );
        native.date_time_format = None;
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(&DataFormat::DateTime(date_time.clone())).unwrap();
        native.suppress_time_format = Some(false);
        assert!(data_format_from_native(&native).is_err());

        let range = UnitRange::new(Unit::Hours, Unit::Milliseconds).unwrap();
        let duration = Duration::custom(Style::Abbreviated, range);
        let mut native = data_format_to_native(&duration.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_DURATION_FORMAT_TYPE));
        assert_eq!(native.duration_style, Some(1));
        assert_eq!(native.duration_unit_largest, Some(4));
        assert_eq!(native.duration_unit_smallest, Some(32));
        assert_eq!(native.use_automatic_duration_units, Some(false));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            DataFormat::Duration(duration)
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

        let checkbox = DataFormat::Checkbox(Checkbox);
        let mut native = data_format_to_native(&checkbox).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_CHECKBOX_FORMAT_TYPE));
        assert_eq!(data_format_from_native(&native).unwrap(), checkbox);
        native.bool_true_string = Some("Yes".to_owned());
        assert!(data_format_from_native(&native).is_err());

        let star_rating = DataFormat::StarRating(StarRating);
        let mut native = data_format_to_native(&star_rating).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_STAR_RATING_FORMAT_TYPE));
        assert_eq!(data_format_from_native(&native).unwrap(), star_rating);
        native.control_maximum = Some(5.0);
        assert!(data_format_from_native(&native).is_err());

        let numeric_control_displays = [
            DisplayFormat::Number(Number::default()),
            DisplayFormat::Currency(Currency::default()),
            DisplayFormat::Percentage(Percentage::default()),
            DisplayFormat::Fraction(Fraction::default()),
            DisplayFormat::Scientific(Scientific::default()),
            DisplayFormat::NumeralSystem(NumeralSystem::default()),
        ];
        for display in numeric_control_displays {
            let native = numeric_control_display_to_native(&display).unwrap();
            assert_eq!(
                numeric_control_display_from_native(&native).unwrap(),
                display
            );
        }
        let invalid_numeric_control_native =
            data_format_to_native(&DataFormat::DateTime(date_time)).unwrap();
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
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::Boolean(true),
        )
        .unwrap();
        editor
            .set_table_cell_checkbox_format(table_id, 1, 1, Checkbox)
            .unwrap();
        editor
            .set_table_cell_checkbox_format(table_id, 1, 2, Checkbox)
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
            Some(Checkbox)
        );
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::Boolean(false))
        );
        assert!(
            reopened
                .reset_table_cell_checkbox_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened.table_cell_checkbox_format(table_id, 1, 2).unwrap(),
            Some(Checkbox)
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
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(3.0).expect("finite test number"),
        )
        .unwrap();
        editor
            .set_table_cell_star_rating_format(table_id, 1, 1, StarRating)
            .unwrap();
        editor
            .set_table_cell_star_rating_format(table_id, 1, 2, StarRating)
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
            Some(StarRating)
        );
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::number(0.0).expect("finite test number"))
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
            Some(StarRating)
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
        let range = Range::new(-10.0, 30.0, 0.5).unwrap();
        let number_display = Number::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        );
        let number_slider = Slider::new(range, number_display.into());
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(25.0).expect("finite test number"),
        )
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
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::number(10.0).expect("finite test number"))
        );

        let currency_slider = Slider::new(range, Currency::default().into());
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
        let range = Range::new(-10.0, 30.0, 0.5).unwrap();
        let number_display = Number::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        );
        let number_stepper = Stepper::new(range, number_display.into());
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(25.0).expect("finite test number"),
        )
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
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::number(-10.0).expect("finite test number"))
        );

        let currency_stepper = Stepper::new(range, Currency::default().into());
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
    fn source_built_table_roundtrips_reuses_and_resets_text_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Text")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::Text("00123".to_owned()),
        )
        .unwrap();
        editor.set_table_cell_text_format(table_id, 1, 1).unwrap();
        editor.set_table_cell_text_format(table_id, 1, 2).unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        assert_eq!(
            formats.entries[0].entry.format,
            Some(text_format_to_native())
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_text_format(table_id, 1, 1).unwrap(),
            Some(Text)
        );
        assert_eq!(
            reopened.table_cell_text_format(table_id, 1, 2).unwrap(),
            Some(Text)
        );
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 1),
            Some(&CellValue::Text("00123".to_owned()))
        );
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::Empty)
        );

        crate::numbers::editor::set_cell_fixture(
            &mut reopened,
            table_id,
            2,
            1,
            CellValue::number(42.0).expect("finite test number"),
        )
        .unwrap();
        let before = reopened.to_bytes().unwrap();
        assert!(reopened.set_table_cell_text_format(table_id, 2, 1).is_err());
        assert_eq!(reopened.to_bytes().unwrap(), before);

        assert!(
            reopened
                .reset_table_cell_text_format(table_id, 1, 1)
                .unwrap()
        );
        assert!(
            reopened
                .reset_table_cell_text_format(table_id, 1, 2)
                .unwrap()
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 2).unwrap();
        let formats = resolve_format_table(reopened.package(), &location).unwrap();
        assert!(formats.entries.is_empty());
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_cleans_up_custom_formats() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Custom Formats")
            .table_dimensions(4, 4)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let name = |value| Name::try_new(value).unwrap();
        let number = CustomNumber::try_with_rules(
            name("Grouped Integer"),
            NumberPattern::try_new("#,###").unwrap(),
            [NumberRule::new(
                Condition::LessThan(ConditionValue::try_new(0.0).unwrap()),
                NumberPattern::try_new("(#,###)").unwrap(),
            )],
        )
        .unwrap();
        let date_time = CustomDateTime::new(
            name("Month Day Year"),
            DateTimePattern::try_new("MMM d, y").unwrap(),
        );
        let text = CustomText::try_new(name("Text With ID Suffix"), "", "ID: ").unwrap();
        let number = Custom::Number(number);
        let date_time = Custom::DateTime(date_time);
        let text = Custom::Text(text);

        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(-12_345.0).expect("finite test number"),
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            2,
            CellValue::number(42.0).expect("finite test number"),
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            2,
            1,
            CellValue::date(789_004_800.0).expect("finite test date"),
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            2,
            2,
            CellValue::Text("Invoice 001".to_owned()),
        )
        .unwrap();
        editor
            .set_table_cell_custom_format(table_id, 1, 1, number.clone())
            .unwrap();
        editor
            .set_table_cell_custom_format(table_id, 1, 2, number.clone())
            .unwrap();
        editor
            .set_table_cell_custom_format(table_id, 2, 1, date_time.clone())
            .unwrap();
        editor
            .set_table_cell_custom_format(table_id, 2, 2, text.clone())
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 3);
        assert_eq!(
            formats
                .entries
                .iter()
                .find(|entry| entry.entry.format.as_ref().unwrap().format_type == Some(270))
                .unwrap()
                .entry
                .refcount,
            2
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_custom_format(table_id, 1, 1).unwrap(),
            Some(number)
        );
        assert_eq!(
            reopened.table_cell_custom_format(table_id, 2, 1).unwrap(),
            Some(date_time)
        );
        assert_eq!(
            reopened.table_cell_custom_format(table_id, 2, 2).unwrap(),
            Some(text)
        );

        for (row, column) in [(1, 1), (1, 2), (2, 1), (2, 2)] {
            assert!(
                reopened
                    .reset_table_cell_custom_format(table_id, row, column)
                    .unwrap()
            );
        }
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 1).unwrap();
        assert!(
            resolve_format_table(reopened.package(), &location)
                .unwrap()
                .entries
                .is_empty()
        );
        assert_eq!(custom::registry_entry_count(reopened.package()).unwrap(), 0);
    }

    #[test]
    fn source_built_table_roundtrips_reuses_and_cleans_up_pop_up_menus() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Menus")
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let priority = PopUpMenu::new(["Low", "Medium", "High"]).unwrap();
        editor
            .set_table_cell_pop_up_menu_format(table_id, 1, 1, priority.clone())
            .unwrap();
        editor
            .set_table_cell_pop_up_menu_format(table_id, 1, 2, priority.clone())
            .unwrap();

        let location = model::locate_attached_cell(editor.package(), table_id, 1, 1).unwrap();
        let formats = resolve_format_table(editor.package(), &location).unwrap();
        assert_eq!(formats.entries.len(), 1);
        assert_eq!(formats.entries[0].entry.refcount, 2);
        assert_eq!(
            formats.entries[0].entry.format,
            Some(text_format_to_native())
        );
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
        let first_model_id = controls.entries[0]
            .entry
            .cell_spec
            .as_ref()
            .unwrap()
            .chooser_control_popup_model
            .as_ref()
            .unwrap()
            .identifier;

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .table_cell_pop_up_menu_format(table_id, 1, 1)
                .unwrap(),
            Some(priority.clone())
        );
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::Text("Low".to_owned()))
        );
        crate::numbers::editor::set_cell_fixture(
            &mut reopened,
            table_id,
            2,
            1,
            CellValue::number(42.0).expect("finite test number"),
        )
        .unwrap();
        let before = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .set_table_cell_pop_up_menu_format(table_id, 2, 1, priority.clone())
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before);

        let blank_priority = PopUpMenu::new(["Low", "Medium", "High"])
            .unwrap()
            .with_initial_selection(InitialSelection::Blank);
        reopened
            .set_table_cell_pop_up_menu_format(table_id, 1, 1, blank_priority.clone())
            .unwrap();
        assert_eq!(
            reopened
                .table_cell_pop_up_menu_format(table_id, 1, 1)
                .unwrap(),
            Some(blank_priority)
        );
        let location = model::locate_attached_cell(reopened.package(), table_id, 1, 1).unwrap();
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &location.object_locations,
            control_table_id,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .unwrap();
        assert_eq!(controls.entries.len(), 2);
        assert!(controls.entries.iter().all(|entry| {
            entry
                .entry
                .cell_spec
                .as_ref()
                .and_then(|spec| spec.chooser_control_popup_model.as_ref())
                .is_some_and(|reference| reference.identifier == first_model_id)
        }));
        assert!(
            reopened
                .reset_table_cell_pop_up_menu_format(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            reopened
                .table_cell_pop_up_menu_format(table_id, 1, 2)
                .unwrap(),
            Some(priority)
        );
        assert!(
            reopened
                .reset_table_cell_pop_up_menu_format(table_id, 1, 2)
                .unwrap()
        );
        let locations = storage::object_locations(reopened.package()).unwrap();
        assert!(!locations.contains_key(&first_model_id));
        let controls = storage::resolve_table_data_list(
            reopened.package(),
            &locations,
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
        let range = UnitRange::hours_to_milliseconds();
        let duration = Duration::custom(Style::Abbreviated, range);
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::duration(3_723.5).expect("finite test duration"),
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            2,
            CellValue::number(1.5).expect("finite test number"),
        )
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
        let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets()[0].tables().next().unwrap().get_cell(1, 2),
            Some(&CellValue::duration(129_600.0).expect("finite test duration"))
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
        let date_time = DateTime::iso_date_time_24_hour_with_seconds();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::date(789_332_889.0).expect("finite test date"),
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            2,
            CellValue::date(789_332_889.0).expect("finite test date"),
        )
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
        let numeral_system = NumeralSystem::new(
            Base::HEXADECIMAL,
            Places::Fixed(FixedPlaces::EIGHT),
            NumeralSystemNegativeStyle::TwosComplement,
        )
        .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite test number"),
        )
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
        let fraction = Fraction::new(FractionAccuracy::Eighths);
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(-12.375).expect("finite test number"),
        )
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
        let scientific = Scientific::new(FixedDecimalPlaces::new(5).unwrap());
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite test number"),
        )
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
        let currency = Currency::new(
            CurrencyCode::USD,
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
            CurrencyStyle::Accounting,
        );
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite test number"),
        )
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
        let format = Number::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        );
        assert_eq!(
            editor.table_cell_number_format(table_id, 1, 1).unwrap(),
            None
        );
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            1,
            CellValue::number(1_234.5).expect("finite test number"),
        )
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
                .set_table_cell_number_format(table_id, 2, 0, Number::default())
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
        let number = Number::new(
            DecimalPlaces::fixed(1).unwrap(),
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        );
        let percentage = Percentage::new(
            DecimalPlaces::fixed(2).unwrap(),
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        );
        editor
            .set_table_cell_data_format(table_id, 1, 1, number.into())
            .unwrap();
        crate::numbers::editor::set_cell_fixture(
            &mut editor,
            table_id,
            1,
            2,
            CellValue::number(-12.345).expect("finite test number"),
        )
        .unwrap();
        assert_eq!(
            editor.table_cell_data_format(table_id, 1, 2).unwrap(),
            DataFormat::Automatic
        );
        editor
            .set_table_cell_percentage_format(table_id, 1, 1, percentage)
            .unwrap();
        editor
            .set_table_cell_percentage_format(table_id, 1, 2, percentage)
            .unwrap();

        assert_eq!(
            editor.table_cell_data_format(table_id, 1, 1).unwrap(),
            DataFormat::Percentage(percentage)
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
            .set_table_cell_data_format(table_id, 1, 1, DataFormat::Automatic)
            .unwrap();
        assert_eq!(
            reopened.table_cell_data_format(table_id, 1, 1).unwrap(),
            DataFormat::Automatic
        );
        assert_eq!(
            reopened
                .table_cell_percentage_format(table_id, 1, 2)
                .unwrap(),
            Some(percentage)
        );
    }
}
