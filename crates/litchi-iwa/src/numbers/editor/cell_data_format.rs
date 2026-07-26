//! Typed data-format storage for native table cells.

use super::*;
use crate::numbers::bnc;
use crate::protobuf::tsk::FormatStructArchive;
#[cfg(test)]
use crate::table_cell_data_format::{
    TableCellCurrencyCode, TableCellCurrencyStyle, TableCellDecimalPlaces,
    TableCellFixedDecimalPlaces, TableCellNegativeNumberStyle, TableCellThousandsSeparator,
};
use crate::table_cell_data_format::{
    TableCellCurrencyFormat, TableCellDataFormat, TableCellFractionFormat, TableCellNumberFormat,
    TableCellPercentageFormat, TableCellScientificFormat,
};

mod codec;
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
            entry
                .entry
                .format
                .as_ref()
                .map(data_format_from_native)
                .transpose()?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers format-table entry {identifier} has no format payload"
                    ))
                })
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Scientific(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Scientific(_) => Err(Error::InvalidFormat(
            "Cannot reset Fraction format from a non-Fraction cell".to_owned(),
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
    set_cell_data_format(package, table_id, row, column, format.into())
}

pub(super) fn set_cell_data_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: TableCellDataFormat,
) -> Result<()> {
    if format == TableCellDataFormat::Automatic {
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
    let cell_format_kind = match format {
        TableCellDataFormat::Currency(_) => bnc::DecimalCellFormatKind::Currency,
        TableCellDataFormat::Number(_)
        | TableCellDataFormat::Percentage(_)
        | TableCellDataFormat::Scientific(_)
        | TableCellDataFormat::Fraction(_) => bnc::DecimalCellFormatKind::NumberOrPercentage,
        TableCellDataFormat::Automatic => unreachable!("handled above"),
    };
    let native = data_format_to_native(format)?;
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

    cell.set_data_format_identifier(new_identifier, cell_format_kind);
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
        | TableCellDataFormat::Fraction(_) => Err(Error::InvalidFormat(
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
    data_format_to_native(format.into()).expect("number format is explicit")
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
    use crate::numbers::{CellValue, NumbersDocumentBuilder, NumbersEditor};
    use crate::table_cell_data_format::{TableCellFractionAccuracy, TableCellFractionFormat};

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
        let native = data_format_to_native(percentage.into()).unwrap();
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
        let mut native = data_format_to_native(currency.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_CURRENCY_FORMAT_TYPE));
        assert_eq!(native.currency_code.as_deref(), Some("EUR"));
        assert_eq!(native.use_accounting_style, Some(true));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Currency(currency)
        );

        native.currency_code = Some("euro".to_owned());
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(currency.into()).unwrap();
        native.use_accounting_style = None;
        assert!(data_format_from_native(&native).is_err());

        let scientific =
            TableCellScientificFormat::new(TableCellFixedDecimalPlaces::new(5).unwrap());
        let mut native = data_format_to_native(scientific.into()).unwrap();
        assert_eq!(native.format_type, Some(NATIVE_SCIENTIFIC_FORMAT_TYPE));
        assert_eq!(native.decimal_places, Some(5));
        assert_eq!(
            data_format_from_native(&native).unwrap(),
            TableCellDataFormat::Scientific(scientific)
        );
        native.negative_style = Some(2);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(scientific.into()).unwrap();
        native.decimal_places = Some(NATIVE_AUTOMATIC_DECIMAL_PLACES);
        assert!(data_format_from_native(&native).is_err());
        native = data_format_to_native(scientific.into()).unwrap();
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
            let native = data_format_to_native(fraction.into()).unwrap();
            assert_eq!(native.format_type, Some(NATIVE_FRACTION_FORMAT_TYPE));
            assert_eq!(native.fraction_accuracy, Some(native_accuracy));
            assert_eq!(
                data_format_from_native(&native).unwrap(),
                TableCellDataFormat::Fraction(fraction)
            );
        }
        let mut invalid = data_format_to_native(TableCellFractionFormat::default().into()).unwrap();
        invalid.fraction_accuracy = Some(3);
        assert!(data_format_from_native(&invalid).is_err());
        invalid = data_format_to_native(TableCellFractionFormat::default().into()).unwrap();
        invalid.decimal_places = Some(2);
        assert!(data_format_from_native(&invalid).is_err());
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
