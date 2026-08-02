//! Typed header and footer configuration for Numbers tables.

use super::*;

mod wire;

use wire::{read_table_header_settings_wire, write_table_header_settings_wire};

const MAX_NATIVE_HEADER_COUNT: u8 = 5;

/// A non-zero Numbers header or footer count accepted by the native formatter.
///
/// Numbers represents a zero count by omitting the corresponding protobuf
/// field. Use `None` in [`NumbersTableHeaderSettings`] to select zero rows or
/// columns.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct NumbersTableHeaderCount(u8);

impl NumbersTableHeaderCount {
    /// One header or footer row or column.
    pub const ONE: Self = Self(1);
    /// Two header or footer rows or columns.
    pub const TWO: Self = Self(2);
    /// Three header or footer rows or columns.
    pub const THREE: Self = Self(3);
    /// Four header or footer rows or columns.
    pub const FOUR: Self = Self(4);
    /// Five header or footer rows or columns.
    pub const FIVE: Self = Self(5);

    /// Validate a native header or footer count.
    pub fn new(count: usize) -> Result<Self> {
        let count = u8::try_from(count).map_err(|_| invalid_header_count())?;
        if !(1..=MAX_NATIVE_HEADER_COUNT).contains(&count) {
            return Err(invalid_header_count());
        }
        Ok(Self(count))
    }

    /// Return the non-zero row or column count.
    pub const fn get(self) -> usize {
        self.0 as usize
    }

    pub(super) fn from_native(count: u32, label: &str) -> Result<Self> {
        Self::new(count as usize).map_err(|_| {
            Error::InvalidFormat(format!(
                "Numbers table {label} count {count} is outside the native 1..=5 range"
            ))
        })
    }

    pub(super) const fn as_native(self) -> u32 {
        self.0 as u32
    }
}

impl TryFrom<usize> for NumbersTableHeaderCount {
    type Error = Error;

    fn try_from(count: usize) -> Result<Self> {
        Self::new(count)
    }
}

impl From<NumbersTableHeaderCount> for usize {
    fn from(count: NumbersTableHeaderCount) -> Self {
        count.get()
    }
}

/// Lossless optional header and footer fields stored by a Numbers table.
///
/// Optional booleans retain their native protobuf presence. This permits a
/// read-modify-write cycle, including clearing a field, without converting an
/// absent default into an explicit value.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct NumbersTableHeaderSettings {
    pub header_rows: Option<NumbersTableHeaderCount>,
    pub header_columns: Option<NumbersTableHeaderCount>,
    pub footer_rows: Option<NumbersTableHeaderCount>,
    pub header_rows_frozen: Option<bool>,
    pub header_columns_frozen: Option<bool>,
    pub repeating_header_rows_enabled: Option<bool>,
    pub repeating_header_columns_enabled: Option<bool>,
}

impl NumbersTableHeaderSettings {
    /// Return the effective number of header rows, treating absence as zero.
    pub fn header_row_count(self) -> usize {
        self.header_rows.map_or(0, NumbersTableHeaderCount::get)
    }

    /// Return the effective number of header columns, treating absence as zero.
    pub fn header_column_count(self) -> usize {
        self.header_columns.map_or(0, NumbersTableHeaderCount::get)
    }

    /// Return the effective number of footer rows, treating absence as zero.
    pub fn footer_row_count(self) -> usize {
        self.footer_rows.map_or(0, NumbersTableHeaderCount::get)
    }

    /// Return whether header rows are effectively frozen.
    pub fn header_rows_are_frozen(self) -> bool {
        self.header_rows_frozen.unwrap_or(false)
    }

    /// Return whether header columns are effectively frozen.
    pub fn header_columns_are_frozen(self) -> bool {
        self.header_columns_frozen.unwrap_or(false)
    }

    /// Return whether header rows effectively repeat when printing.
    pub fn repeats_header_rows(self) -> bool {
        self.repeating_header_rows_enabled.unwrap_or(false)
    }

    /// Return whether header columns effectively repeat when printing.
    pub fn repeats_header_columns(self) -> bool {
        self.repeating_header_columns_enabled.unwrap_or(false)
    }

    pub(super) fn from_model(model: &TableModelArchive) -> Result<Self> {
        Ok(Self {
            header_rows: model
                .number_of_header_rows
                .map(|count| NumbersTableHeaderCount::from_native(count, "header row"))
                .transpose()?,
            header_columns: model
                .number_of_header_columns
                .map(|count| NumbersTableHeaderCount::from_native(count, "header column"))
                .transpose()?,
            footer_rows: model
                .number_of_footer_rows
                .map(|count| NumbersTableHeaderCount::from_native(count, "footer row"))
                .transpose()?,
            header_rows_frozen: model.header_rows_frozen,
            header_columns_frozen: model.header_columns_frozen,
            repeating_header_rows_enabled: model.repeating_header_rows_enabled,
            repeating_header_columns_enabled: model.repeating_header_columns_enabled,
        })
    }
}

impl NumbersEditor {
    /// Read the lossless header and footer configuration of an attached table.
    pub fn table_header_settings(&self, table_id: u64) -> Result<NumbersTableHeaderSettings> {
        read_attached_table_header_settings(&self.package, table_id)
    }

    /// Replace an attached table's header and footer configuration transactionally.
    pub fn set_table_header_settings(
        &mut self,
        table_id: u64,
        settings: NumbersTableHeaderSettings,
    ) -> Result<()> {
        if read_attached_table_header_settings(&self.package, table_id)? == settings {
            return Ok(());
        }
        let mut staged = self.package.clone();
        set_attached_table_header_settings(&mut staged, table_id, settings)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.table_header_settings(table_id)? != settings {
            return Err(Error::InvalidFormat(
                "Numbers table header settings failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }
}

fn invalid_header_count() -> Error {
    Error::ParseError(format!(
        "Numbers table header and footer counts must be in 1..={MAX_NATIVE_HEADER_COUNT}"
    ))
}

fn read_table_header_settings(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<NumbersTableHeaderSettings> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let message_index = table_model_message_index(object, descriptor.object_id)?;
    let settings = read_table_header_settings_wire(
        object.messages[message_index].data.as_slice(),
        &descriptor.model,
    )?;
    validate_table_header_settings(&descriptor.model, settings).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers table model {} has invalid stored header settings: {error}",
            descriptor.object_id
        ))
    })?;
    Ok(settings)
}

pub(super) fn read_attached_table_header_settings(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<NumbersTableHeaderSettings> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    read_table_header_settings(package, &descriptor)
}

pub(super) fn set_attached_table_header_settings(
    package: &mut IWorkPackage,
    table_id: u64,
    settings: NumbersTableHeaderSettings,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_table_header_settings(&descriptor.model, settings)?;
    if read_table_header_settings(package, &descriptor)? == settings {
        return Ok(());
    }
    let locations = object_locations(package)?;
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
        })?;
        let message_index = table_model_message_index(object, table_id)?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let model = TableModelArchive::decode(original)?;
        let data = write_table_header_settings_wire(original, &model, settings)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn table_model_message_index(object: &ArchiveObject, table_id: u64) -> Result<usize> {
    let matches = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            matches!(message.type_, 6000 | 6001)
                .then(|| {
                    TableModelArchive::decode(message.data.as_slice())
                        .ok()
                        .map(|_| index)
                })
                .flatten()
        })
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [index] => Ok(*index),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers table model object {table_id} must contain exactly one decodable payload, found {}",
            matches.len()
        ))),
    }
}

fn validate_table_header_settings(
    model: &TableModelArchive,
    settings: NumbersTableHeaderSettings,
) -> Result<()> {
    let header_rows = settings.header_rows.map_or(0, NumbersTableHeaderCount::get);
    let footer_rows = settings.footer_rows.map_or(0, NumbersTableHeaderCount::get);
    let rows = model.number_of_rows as usize;
    if header_rows + footer_rows > rows {
        return Err(Error::ParseError(format!(
            "Numbers header rows ({header_rows}) plus footer rows ({footer_rows}) exceed the table's {rows} rows"
        )));
    }
    let header_columns = settings
        .header_columns
        .map_or(0, NumbersTableHeaderCount::get);
    let columns = model.number_of_columns as usize;
    if header_columns > columns {
        return Err(Error::ParseError(format!(
            "Numbers header columns ({header_columns}) exceed the table's {columns} columns"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn header_count_accepts_only_native_non_zero_range() {
        assert!(NumbersTableHeaderCount::new(0).is_err());
        for count in 1..=5 {
            assert_eq!(NumbersTableHeaderCount::new(count).unwrap().get(), count);
        }
        assert!(NumbersTableHeaderCount::new(6).is_err());
        assert!(NumbersTableHeaderCount::new(usize::MAX).is_err());
    }
}
