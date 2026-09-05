//! Compatibility reads for native iWork table lock state.

use crate::wire::parse_wire_fields;
use crate::{Error, Result};
use litchi_iwa_common::table::lock::State as TableLockState;

const TABLE_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;

/// Read effective lock state directly from a `TST.TableInfoArchive` payload.
pub(crate) fn table_lock_state_from_message(data: &[u8]) -> Result<TableLockState> {
    Ok(TableLockState::from_locked(
        raw_table_lock_state(data)?.unwrap_or(false),
    ))
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
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(
        &data[field.payload_start()..field.end()],
    )
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
    use crate::wire::{append_varint_field, transform_length_delimited_field};

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
    fn missing_lock_is_unlocked() {
        let source = table_info(None);
        assert_eq!(
            table_lock_state_from_message(&source).unwrap(),
            TableLockState::Unlocked
        );
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
