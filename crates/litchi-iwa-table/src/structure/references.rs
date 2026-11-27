//! Raw `TSP.Reference` lists and their archive-header metadata.
//!
//! Payload edits stay at the wire level: a repeated reference field is
//! rewritten by reusing each retained record's exact bytes and encoding only
//! new records, so unknown fields and non-canonical encodings of untouched
//! records survive. The archive header's reference metadata follows through
//! the core's exact transition primitive.

use litchi_iwa_common::wire::{
    parse_wire_view, repeated_length_delimited_payloads, repeated_varint_values,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{ArchiveObject, RawMessage};

use super::{Fault, FaultResult, copy_bytes, map_core, map_wire, reserve};

/// `TSP.Reference.identifier`.
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// Encode a canonical `TSP.Reference` payload for `identifier`.
pub fn reference_payload(identifier: u64) -> FaultResult<Vec<u8>> {
    let mut output = Vec::new();
    reserve(&mut output, 11)?;
    output.push(0x08);
    let mut value = identifier;
    loop {
        let byte = u8::try_from(value & 0x7f).map_err(|_error| Fault::Invalid)?;
        value >>= 7;
        if value == 0 {
            output.push(byte);
            break;
        }
        output.push(byte | 0x80);
    }
    Ok(output)
}

/// The identifier of a `TSP.Reference` payload. Records carrying anything but
/// their identifier are refused: a structural edit never reinterprets them.
pub fn reference_identifier(payload: &[u8]) -> FaultResult<u64> {
    let view = parse_wire_view(payload, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|error| map_wire(&error))?;
    if view
        .fields()
        .any(|field| field.number() != REFERENCE_IDENTIFIER_FIELD || field.wire_type() != 0)
    {
        return Err(Fault::Unsupported);
    }
    let values = repeated_varint_values(payload, REFERENCE_IDENTIFIER_FIELD)
        .map_err(|error| map_wire(&error))?;
    match values.as_slice() {
        [identifier] if *identifier != 0 => Ok(*identifier),
        _ => Err(Fault::Invalid),
    }
}

/// The raw records of the repeated reference field at `path` of `data`.
pub fn records<'d>(data: &'d [u8], path: &[u32]) -> FaultResult<Vec<&'d [u8]>> {
    match path {
        [field] => {
            repeated_length_delimited_payloads(data, *field).map_err(|error| map_wire(&error))
        },
        [outer, rest @ ..] => {
            let payloads = repeated_length_delimited_payloads(data, *outer)
                .map_err(|error| map_wire(&error))?;
            let [inner] = payloads.as_slice() else {
                return Err(Fault::Invalid);
            };
            records(inner, rest)
        },
        [] => Err(Fault::Invalid),
    }
}

/// The identifiers of the repeated reference field at `path` of `data`.
pub fn identifiers(data: &[u8], path: &[u32]) -> FaultResult<Vec<u64>> {
    let records = records(data, path)?;
    let mut output = Vec::new();
    reserve(&mut output, records.len())?;
    for record in records {
        output.push(reference_identifier(record)?);
    }
    Ok(output)
}

/// Rewrite the repeated reference field at `path` to `after`, reusing the
/// exact bytes of every retained record and encoding new ones canonically.
pub fn rewrite(data: &[u8], path: &[u32], after: &[u64]) -> FaultResult<Vec<u8>> {
    match path {
        [field] => {
            let records = records(data, path)?;
            let mut before = Vec::new();
            reserve(&mut before, records.len())?;
            for record in &records {
                before.push(reference_identifier(record)?);
            }
            let mut replacements = Vec::new();
            reserve(&mut replacements, after.len())?;
            for identifier in after {
                if after.iter().filter(|other| *other == identifier).count() != 1 {
                    return Err(Fault::Invalid);
                }
                match before.iter().position(|existing| existing == identifier) {
                    Some(index) => replacements.push(copy_bytes(records[index])?),
                    None => replacements.push(reference_payload(*identifier)?),
                }
            }
            rewrite_repeated_length_delimited_fields(data, *field, &replacements)
                .map_err(|error| map_wire(&error))
        },
        [outer, rest @ ..] => transform_length_delimited_field(data, *outer, |inner| {
            rewrite(inner, rest, after).map_err(WireFault)
        })
        .map_err(|error: WireFault| error.0),
        [] => Err(Fault::Invalid),
    }
}

/// A [`Fault`] carried through the common wire transform callbacks.
struct WireFault(Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(map_wire(&error))
    }
}

/// Replace message `index` of `object` with `data`, transitioning its header
/// references: `removed` leave the aggregate list and every field list,
/// `appended` join the aggregate list as its suffix and the field list at
/// `field_path` when the header describes that field.
pub fn replace(
    object: &mut ArchiveObject,
    index: usize,
    data: Vec<u8>,
    removed: &[u64],
    appended: &[u64],
    field_path: &[u32],
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Fault::Invalid)?;
    let message_type = object.messages.get(index).ok_or(Fault::Invalid)?.type_;
    let aggregate_before = copy_slice(&info.object_references)?;
    let mut aggregate_after = Vec::new();
    reserve(
        &mut aggregate_after,
        aggregate_before.len() + appended.len(),
    )?;
    aggregate_after.extend(
        aggregate_before
            .iter()
            .copied()
            .filter(|identifier| !removed.contains(identifier)),
    );
    for identifier in appended {
        if !aggregate_after.contains(identifier) {
            aggregate_after.push(*identifier);
        }
    }
    let mut owned = Vec::new();
    for (field_info_index, field) in info.field_infos.iter().enumerate() {
        let touches_removed = field
            .object_references
            .iter()
            .any(|identifier| removed.contains(identifier));
        let takes_appended = !appended.is_empty() && field.path.as_slice() == field_path;
        if !touches_removed && !takes_appended {
            continue;
        }
        let before = copy_slice(&field.object_references)?;
        let mut after = Vec::new();
        reserve(&mut after, before.len() + appended.len())?;
        after.extend(
            before
                .iter()
                .copied()
                .filter(|identifier| !removed.contains(identifier)),
        );
        if takes_appended {
            for identifier in appended {
                if !after.contains(identifier) {
                    after.push(*identifier);
                }
            }
        }
        reserve(&mut owned, 1)?;
        owned.push((
            field_info_index,
            copy_slice(field.path.as_slice())?,
            before,
            after,
        ));
    }
    let fields: Vec<FieldObjectReferenceTransition<'_>> = owned
        .iter()
        .map(
            |(field_info_index, path, before, after)| FieldObjectReferenceTransition {
                field_info_index: *field_info_index,
                expected_path: path,
                before,
                after,
            },
        )
        .collect();
    object
        .replace_message_transitioning_object_references_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: &aggregate_before,
                aggregate_after: &aggregate_after,
                fields: &fields,
            },
            limits,
        )
        .map_err(|error| map_core(&error))?;
    Ok(())
}

fn copy_slice<T: Copy>(source: &[T]) -> FaultResult<Vec<T>> {
    let mut output = Vec::new();
    reserve(&mut output, source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}
