//! Decimal-number format storage for native table cells.

use super::*;
use crate::numbers::bnc;
use crate::protobuf::tsk::FormatStructArchive;
use crate::table_cell_number_format::{
    TableCellDecimalPlaces, TableCellFixedDecimalPlaces, TableCellNegativeNumberStyle,
    TableCellNumberFormat, TableCellThousandsSeparator,
};

const EXPLICIT_NUMBER_FORMAT: u16 = 1;
const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;
const NATIVE_AUTOMATIC_DECIMAL_PLACES: u32 = 253;

pub(super) fn cell_number_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellNumberFormat>> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(None);
    };
    let cell = BncCell::parse(&data)?;
    match format_reference(&cell)? {
        CellFormatReference::Automatic { .. } => Ok(None),
        CellFormatReference::Number(identifier) => {
            let resolved = resolve_format_table(package, &location)?;
            let entry = required_format_entry(&resolved, identifier)?;
            entry
                .entry
                .format
                .as_ref()
                .map(number_format_from_native)
                .transpose()
        },
    }
}

pub(super) fn set_cell_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    format: TableCellNumberFormat,
) -> Result<()> {
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
    let old_identifier = old_reference.identifier();
    let native = number_format_to_native(format);
    let resolved = resolve_format_table(package, &location)?;
    let old_entry = old_identifier
        .map(|identifier| required_format_entry(&resolved, identifier))
        .transpose()?;

    if matches!(old_reference, CellFormatReference::Number(_))
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
        if old_identifier != Some(reusable.entry.key) {
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

    cell.set_number_format_identifier(new_identifier);
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

    if let Some(old_entry) = old_entry
        && old_entry.entry.key != new_identifier
    {
        storage::decrement_table_data_list_entry(
            package,
            &location.object_locations,
            &resolved,
            old_entry,
            tst::table_data_list::ListType::Format,
        )?;
    }
    Ok(())
}

pub(super) fn reset_cell_number_format(
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
    let CellFormatReference::Number(identifier) = format_reference(&cell)? else {
        return Ok(false);
    };
    let resolved = resolve_format_table(package, &location)?;
    let old_entry = required_format_entry(&resolved, identifier)?;

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
    storage::decrement_table_data_list_entry(
        package,
        &location.object_locations,
        &resolved,
        old_entry,
        tst::table_data_list::ListType::Format,
    )?;
    Ok(true)
}

#[derive(Clone, Copy)]
enum CellFormatReference {
    Automatic { identifier: Option<u32> },
    Number(u32),
}

impl CellFormatReference {
    const fn identifier(self) -> Option<u32> {
        match self {
            Self::Automatic { identifier } => identifier,
            Self::Number(identifier) => Some(identifier),
        }
    }
}

fn format_reference(cell: &BncCell) -> Result<CellFormatReference> {
    let explicit = cell.explicit_format_flags();
    let kind = cell.cell_format_kind();
    let identifier = cell.format_identifier();
    match (explicit, kind, identifier) {
        (0, None, None) => Ok(CellFormatReference::Automatic { identifier: None }),
        (0, Some(bnc::NUMBER_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Automatic {
                identifier: Some(identifier),
            })
        },
        (EXPLICIT_NUMBER_FORMAT, Some(bnc::NUMBER_CELL_FORMAT_KIND), Some(identifier)) => {
            Ok(CellFormatReference::Number(identifier))
        },
        (EXPLICIT_NUMBER_FORMAT, Some(kind), _) => Err(Error::InvalidFormat(format!(
            "Table cell uses unsupported explicit format kind {kind}"
        ))),
        _ => Err(Error::InvalidFormat(
            "Table cell contains inconsistent number-format metadata".to_owned(),
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

fn number_format_to_native(format: TableCellNumberFormat) -> FormatStructArchive {
    FormatStructArchive {
        format_type: Some(NATIVE_NUMBER_FORMAT_TYPE),
        decimal_places: Some(match format.decimal_places() {
            TableCellDecimalPlaces::Automatic => NATIVE_AUTOMATIC_DECIMAL_PLACES,
            TableCellDecimalPlaces::Fixed(places) => u32::from(places.value()),
        }),
        negative_style: Some(match format.negative_style() {
            TableCellNegativeNumberStyle::MinusSign => 0,
            TableCellNegativeNumberStyle::Red => 1,
            TableCellNegativeNumberStyle::Parentheses => 2,
            TableCellNegativeNumberStyle::RedParentheses => 3,
        }),
        show_thousands_separator: Some(matches!(
            format.thousands_separator(),
            TableCellThousandsSeparator::Shown
        )),
        ..Default::default()
    }
}

fn number_format_from_native(native: &FormatStructArchive) -> Result<TableCellNumberFormat> {
    if native.format_type != Some(NATIVE_NUMBER_FORMAT_TYPE) {
        return Err(Error::InvalidFormat(format!(
            "Table cell uses unsupported native number-format type {:?}",
            native.format_type
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
                "Table-cell number format has no decimal-place setting".to_owned(),
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
                "Table-cell number format has no thousands-separator setting".to_owned(),
            ));
        },
    };
    Ok(TableCellNumberFormat::new(
        decimal_places,
        negative_style,
        thousands_separator,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::{CellValue, NumbersDocumentBuilder, NumbersEditor};

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
}
