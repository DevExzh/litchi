//! Private native owner for focused decimal display-format transactions.
//!
//! Number and Percentage cells use the same BNC decimal cell kind and the same
//! format-list graph.  This module keeps that graph surgery in one route while
//! retaining a typed family boundary at every codec and semantic conversion.
//! It intentionally exposes no native identifiers or arbitrary format-type
//! input to the package API.

use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_number_format_codec as number_codec,
    numbers_table_cell_percentage_format_codec as percentage_codec,
    numbers_table_cell_storage_codec as storage_codec,
};
use litchi_numbers_wire::BncCell;

use super::table_cell_control_native as native;
use super::{
    Package, table_cell_pop_up_menu as popup,
    table_cell_pop_up_menu::{CellTarget, Error, Path, TransactionBudget},
    table_cell_pop_up_menu_native as popup_native,
};
use crate::cell::data_format::number::{
    DecimalPlaces, FixedDecimalPlaces, NegativeStyle, Number, Percentage, ThousandsSeparator,
};

const NUMBER_FORMAT_TYPE: u32 = number_codec::NATIVE_NUMBER_FORMAT_TYPE;
const PERCENTAGE_FORMAT_TYPE: u32 = percentage_codec::NATIVE_PERCENTAGE_FORMAT_TYPE;

/// The only display families admitted by the focused decimal owner.
///
/// The enum is private to the package adapter.  In particular, callers cannot
/// supply an arbitrary native discriminator and use this route as a generic
/// raw-ID editor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DisplayFormatFamily {
    Number,
    Percentage,
}

impl DisplayFormatFamily {
    const fn native_type(self) -> u32 {
        match self {
            Self::Number => NUMBER_FORMAT_TYPE,
            Self::Percentage => PERCENTAGE_FORMAT_TYPE,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NativeDisplayValue {
    Number(Number),
    Percentage(Percentage),
}

impl NativeDisplayValue {
    const fn family(self) -> DisplayFormatFamily {
        match self {
            Self::Number(_) => DisplayFormatFamily::Number,
            Self::Percentage(_) => DisplayFormatFamily::Percentage,
        }
    }

    const fn decimal_parts(self) -> (DecimalPlaces, NegativeStyle, ThousandsSeparator) {
        match self {
            Self::Number(value) => (
                value.decimal_places(),
                value.negative_style(),
                value.thousands_separator(),
            ),
            Self::Percentage(value) => (
                value.decimal_places(),
                value.negative_style(),
                value.thousands_separator(),
            ),
        }
    }

    fn from_parts(
        family: DisplayFormatFamily,
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
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
            DisplayFormatFamily::Percentage => NativeDisplayValue::Percentage(Percentage::new(
                decimal_places,
                negative_style,
                thousands_separator,
            )),
        };
        Ok(value)
    }
}

fn rewrite_display_cell_metadata(
    source: &[u8],
    desired_identifier: Option<u32>,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let mut cell = BncCell::parse(source).map_err(|_| Error::InvalidSource { path })?;
    let source_value = cell.stored_value();
    let source_cache = cell
        .cached_scalar()
        .map_err(|_| Error::InvalidSource { path })?;
    cell.set_number_or_percentage_format_identifier_preserving_value(desired_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let output_limit = source
        .len()
        .checked_add(8)
        .ok_or(Error::InvalidSource { path })?;
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
        &candidate_cell,
        desired_identifier,
    )?;
    Ok(output)
}

fn verify_display_cell_metadata(
    source_value: litchi_numbers_wire::StoredValue,
    source_cache: Option<litchi_numbers_wire::CachedScalar>,
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
        Ok(Some(NativeDisplayValue::Percentage(_))) | Err(DisplayReadError::WrongFormatFamily) => {
            Err(NumberFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(NumberFormatReadError::Native(error)),
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
        Ok(Some(NativeDisplayValue::Number(_))) | Err(DisplayReadError::WrongFormatFamily) => {
            Err(PercentageFormatReadError::WrongFormatFamily)
        },
        Err(DisplayReadError::Native(error)) => Err(PercentageFormatReadError::Native(error)),
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
    let cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
    let format_identifier = cell.format_identifier();
    if cell.control_cell_spec_identifier().is_some() {
        return Err(DisplayReadError::WrongFormatFamily);
    }
    match cell.cell_format_kind() {
        Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND) => {},
        Some(_) => return Err(DisplayReadError::WrongFormatFamily),
        None if format_identifier.is_some() || cell.explicit_format_flags() != 0 => {
            return Err(DisplayReadError::WrongFormatFamily);
        },
        None => return Ok(None),
    }
    let format_identifier = format_identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource { path })?;
    if cell.secondary_format_identifier().is_some() {
        return Err(DisplayReadError::WrongFormatFamily);
    }
    let explicit_flags = cell.explicit_format_flags();
    if explicit_flags != 0 && explicit_flags != litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT {
        return Err(DisplayReadError::WrongFormatFamily);
    }

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
    let cell = BncCell::parse(cell_source).map_err(|_| Error::InvalidSource { path })?;
    let old_format = cell.format_identifier();
    if cell.control_cell_spec_identifier().is_some() {
        return Err(Error::UnsupportedDependency { path });
    }
    match (old_format, cell.cell_format_kind()) {
        (Some(identifier), Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND))
            if identifier != 0 => {},
        (None, None) => {},
        _ => return Err(Error::UnsupportedDependency { path }),
    }
    if old_format.is_some() && cell.secondary_format_identifier().is_some() {
        return Err(Error::UnsupportedDependency { path });
    }
    let explicit_flags = cell.explicit_format_flags();
    if explicit_flags != 0 && explicit_flags != litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT {
        return Err(Error::UnsupportedDependency { path });
    }
    if old_format.is_none() && explicit_flags != 0 {
        return Err(Error::InvalidSource { path });
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
        if explicit_flags == litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT {
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
    let (new_format, new_format_key) = native::mutate_list(
        &format_payload,
        old_format,
        desired_payload.as_deref(),
        true,
        budget,
        path,
    )?;
    let replacement_cell = if after.is_some() {
        let key = new_format_key.ok_or(Error::InvalidSource { path })?;
        rewrite_display_cell_metadata(cell_source, Some(key), path)?
    } else {
        rewrite_display_cell_metadata(cell_source, None, path)?
    };
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
                path,
            )?;
            Ok(value)
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
    // failure, so successful Number/Percentage reads retain their existing
    // resource accounting. A malformed payload remains a native error.
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
    let (decimal_places, negative_style, thousands_separator) = value.decimal_parts();
    let options = native::control_codec_options(source.len(), budget);
    match family {
        DisplayFormatFamily::Number => {
            let NativeDisplayValue::Number(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = number_codec::prepare_number_format_rewrite(
                source,
                number_codec::NumberFormatWrite::new(
                    native_decimal_places(decimal_places),
                    native_negative_style(negative_style),
                    matches!(thousands_separator, ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_number_rewrite(prepared, budget, path)
        },
        DisplayFormatFamily::Percentage => {
            let NativeDisplayValue::Percentage(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = percentage_codec::prepare_percentage_format_rewrite(
                source,
                percentage_codec::PercentageFormatWrite::new(
                    native_decimal_places(decimal_places),
                    native_negative_style(negative_style),
                    matches!(thousands_separator, ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_percentage_rewrite(prepared, budget, path)
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
    let (decimal_places, negative_style, thousands_separator) = value.decimal_parts();
    let options = native::control_codec_options(source_len, budget);
    match family {
        DisplayFormatFamily::Number => {
            let NativeDisplayValue::Number(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = number_codec::prepare_number_format_write(
                number_codec::NumberFormatWrite::new(
                    native_decimal_places(decimal_places),
                    native_negative_style(negative_style),
                    matches!(thousands_separator, ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_number_append(prepared, budget, path)
        },
        DisplayFormatFamily::Percentage => {
            let NativeDisplayValue::Percentage(_) = value else {
                return Err(Error::UnsupportedDependency { path });
            };
            let prepared = percentage_codec::prepare_percentage_format_write(
                percentage_codec::PercentageFormatWrite::new(
                    native_decimal_places(decimal_places),
                    native_negative_style(negative_style),
                    matches!(thousands_separator, ThousandsSeparator::Shown),
                ),
                options,
            )
            .map_err(|error| native::map_control_error(error, path))?;
            execute_percentage_append(prepared, budget, path)
        },
    }
}

fn native_decimal_places(value: DecimalPlaces) -> u32 {
    match value {
        DecimalPlaces::Automatic => 253,
        DecimalPlaces::Fixed(value) => u32::from(value.value()),
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

fn display_read_error_to_write_error(error: DisplayReadError, path: Path) -> Error {
    match error {
        DisplayReadError::WrongFormatFamily => Error::UnsupportedDependency { path },
        DisplayReadError::Native(error) => error,
    }
}
