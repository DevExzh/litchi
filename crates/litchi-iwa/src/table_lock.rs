//! Wire-preserving lock state for native iWork table drawables.

use crate::archive::RawMessage;
use crate::protobuf::tst::TableInfoArchive;
use crate::wire::{parse_wire_fields, patch_varint_field, transform_length_delimited_field};
use crate::{Error, IWorkPackage, Result};
use prost::Message;

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;

/// Interactive editing state of a native iWork table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TableLockState {
    /// The table can be selected and edited interactively.
    #[default]
    Unlocked,
    /// The table is protected from interactive editing.
    Locked,
}

impl TableLockState {
    /// Construct a lock state from its native boolean representation.
    pub const fn from_locked(locked: bool) -> Self {
        if locked { Self::Locked } else { Self::Unlocked }
    }

    /// Return whether the table is protected from interactive editing.
    pub const fn is_locked(self) -> bool {
        matches!(self, Self::Locked)
    }
}

/// Read one table's effective interactive lock state.
pub(crate) fn table_lock_state(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    application: &str,
) -> Result<TableLockState> {
    let (_, _, message) =
        table_message(package, archive_name, drawable_object_id, application, None)?;
    table_lock_state_from_message(&message)
}

/// Read a Numbers table whose table-info type varies by archive generation.
pub(crate) fn table_lock_state_for_model(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    model_object_id: u64,
) -> Result<TableLockState> {
    let (_, _, message) = table_message(
        package,
        archive_name,
        drawable_object_id,
        "Numbers",
        Some(model_object_id),
    )?;
    table_lock_state_from_message(&message)
}

/// Read effective lock state directly from a `TST.TableInfoArchive` payload.
pub(crate) fn table_lock_state_from_message(data: &[u8]) -> Result<TableLockState> {
    Ok(TableLockState::from_locked(
        raw_table_lock_state(data)?.unwrap_or(false),
    ))
}

/// Set one table's lock state without normalizing unrelated protobuf bytes.
pub(crate) fn set_table_lock_state(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    application: &str,
    state: TableLockState,
) -> Result<()> {
    set_table_lock_state_inner(
        package,
        archive_name,
        drawable_object_id,
        application,
        None,
        state,
    )
}

/// Set a Numbers table whose table-info type varies by archive generation.
pub(crate) fn set_table_lock_state_for_model(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    model_object_id: u64,
    state: TableLockState,
) -> Result<()> {
    set_table_lock_state_inner(
        package,
        archive_name,
        drawable_object_id,
        "Numbers",
        Some(model_object_id),
        state,
    )
}

fn set_table_lock_state_inner(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    application: &str,
    model_object_id: Option<u64>,
    state: TableLockState,
) -> Result<()> {
    let (message_index, message_type, message) = table_message(
        package,
        archive_name,
        drawable_object_id,
        application,
        model_object_id,
    )?;
    let current = raw_table_lock_state(&message)?;
    if current.unwrap_or(false) == state.is_locked() {
        return Ok(());
    }
    let data =
        transform_length_delimited_field(&message, TABLE_DRAWABLE_SUPER_FIELD, |drawable| {
            patch_varint_field(
                drawable,
                DRAWABLE_LOCKED_FIELD,
                current.is_some(),
                replacement_presence(current, state.is_locked()).map(u64::from),
            )
        })?;
    if table_lock_state_from_message(&data)? != state {
        return Err(Error::InvalidFormat(format!(
            "{application} table {drawable_object_id} lock update failed validation"
        )));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{application} table drawable {drawable_object_id} is missing"
            ))
        })?;
        Ok(object
            .replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )
            .map(|_| ())?)
    })
}

const fn replacement_presence(current: Option<bool>, replacement: bool) -> Option<bool> {
    if current.is_some() || replacement {
        Some(replacement)
    } else {
        None
    }
}

fn raw_table_lock_state(data: &[u8]) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let drawable = singular_field(&fields, TABLE_DRAWABLE_SUPER_FIELD, "table drawable super")?;
    require_wire_type(drawable, 2, "table drawable super")?;
    strict_optional_bool(
        &data[drawable.payload_start()..drawable.end()],
        DRAWABLE_LOCKED_FIELD,
        "table lock",
    )
}

fn table_message(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_object_id: u64,
    application: &str,
    model_object_id: Option<u64>,
) -> Result<(usize, u32, Vec<u8>)> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{application} table drawable {drawable_object_id} is missing"
        ))
    })?;
    let mut messages = object.messages.iter().enumerate().filter(|(_, message)| {
        if let Some(model_object_id) = model_object_id {
            TableInfoArchive::decode(message.data.as_slice())
                .is_ok_and(|info| info.table_model.identifier == model_object_id)
        } else {
            message.type_ == TABLE_INFO_MESSAGE_TYPE
        }
    });
    let Some((message_index, message)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{application} drawable {drawable_object_id} has no table-info payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{application} drawable {drawable_object_id} has multiple table-info payloads"
        )));
    }
    Ok((message_index, message.type_, message.data.clone()))
}

fn strict_optional_bool(data: &[u8], field_number: u32, label: &str) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{label} field occurs more than once"
        )));
    }
    require_wire_type(field, 0, label)?;
    let (value, length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.payload_start()..field.end()])
            .map_err(|error| Error::InvalidFormat(format!("invalid {label}: {error}")))?;
    if field.payload_start() + length != field.end() {
        return Err(Error::InvalidFormat(format!(
            "{label} contains trailing bytes"
        )));
    }
    match value {
        0 => Ok(Some(false)),
        1 => Ok(Some(true)),
        _ => Err(Error::InvalidFormat(format!(
            "{label} must be encoded as zero or one, found {value}"
        ))),
    }
}

fn singular_field<'a>(
    fields: &'a [crate::wire::WireField],
    field_number: u32,
    label: &str,
) -> Result<&'a crate::wire::WireField> {
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Err(Error::InvalidFormat(format!(
            "{label} must occur exactly once, found none"
        )));
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{label} must occur exactly once, found multiple"
        )));
    }
    Ok(field)
}

fn require_wire_type(field: &crate::wire::WireField, expected: u8, label: &str) -> Result<()> {
    if field.wire_type() != expected {
        return Err(Error::InvalidFormat(format!(
            "{label} has wire type {}, expected {expected}",
            field.wire_type()
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use prost::Message;

    use super::*;
    use crate::protobuf::{tsd, tsp, tst};
    use crate::wire::append_varint_field;

    fn table_info(locked: Option<bool>) -> Vec<u8> {
        tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                locked,
                ..Default::default()
            },
            table_model: tsp::Reference {
                identifier: 42,
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec()
    }

    #[test]
    fn lock_patch_preserves_unknowns_and_explicit_defaults() {
        let mut source = table_info(Some(false));
        append_varint_field(&mut source, 200, 17).unwrap();
        let current = raw_table_lock_state(&source).unwrap();
        let locked =
            transform_length_delimited_field(&source, TABLE_DRAWABLE_SUPER_FIELD, |drawable| {
                patch_varint_field(drawable, DRAWABLE_LOCKED_FIELD, current.is_some(), Some(1))
            })
            .unwrap();
        assert_eq!(
            table_lock_state_from_message(&locked).unwrap(),
            TableLockState::Locked
        );
        let restored =
            transform_length_delimited_field(&locked, TABLE_DRAWABLE_SUPER_FIELD, |drawable| {
                patch_varint_field(drawable, DRAWABLE_LOCKED_FIELD, true, Some(0))
            })
            .unwrap();
        assert_eq!(restored, source);
    }

    #[test]
    fn missing_lock_stays_absent_when_effectively_unchanged() {
        let source = table_info(None);
        assert_eq!(
            table_lock_state_from_message(&source).unwrap(),
            TableLockState::Unlocked
        );
        assert_eq!(replacement_presence(None, false), None);
        assert_eq!(replacement_presence(None, true), Some(true));
    }

    #[test]
    fn rejects_duplicate_and_non_boolean_lock_fields() {
        let duplicate = transform_length_delimited_field(&table_info(Some(false)), 1, |drawable| {
            let mut data = drawable.to_vec();
            append_varint_field(&mut data, DRAWABLE_LOCKED_FIELD, 1)?;
            Ok(data)
        })
        .unwrap();
        assert!(table_lock_state_from_message(&duplicate).is_err());

        let invalid = transform_length_delimited_field(&table_info(None), 1, |drawable| {
            let mut data = drawable.to_vec();
            append_varint_field(&mut data, DRAWABLE_LOCKED_FIELD, 2)?;
            Ok(data)
        })
        .unwrap();
        assert!(table_lock_state_from_message(&invalid).is_err());
    }
}
