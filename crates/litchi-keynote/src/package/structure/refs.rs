//! Schema-directed object and data references inside native message payloads.
//!
//! The tables in [`super::schema`] say where each admitted message type keeps
//! its `TSP.Reference` and `TSP.DataReference` values. [`visit`] reports every
//! reference of one payload; [`rewrite`] re-encodes a payload with remapped
//! identifiers, copying every other byte (unknown fields included) verbatim.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_common::wire::{WireField, parse_wire_fields};

use super::schema::{Field, MESSAGES, Slot, root};

/// Maximum nesting followed below one message payload.
const MAX_DEPTH: usize = 64;
const WIRE_VARINT: u8 = 0;
const WIRE_LENGTH_DELIMITED: u8 = 2;

/// What a reference names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(in crate::package) enum RefKind {
    /// A native object identifier.
    Object,
    /// A package data identifier.
    Data,
}

/// Why a payload could not be walked.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) enum RefError {
    /// The message type has no admitted reference schema.
    UnsupportedType(u32),
    /// The payload is not well-formed protobuf for its schema.
    Malformed,
    /// The payload nests deeper than [`MAX_DEPTH`].
    Depth,
}

impl From<litchi_iwa_common::Error> for RefError {
    fn from(_error: litchi_iwa_common::Error) -> Self {
        Self::Malformed
    }
}

/// Whether `message_type` has an admitted reference schema.
pub(in crate::package) const fn admitted(message_type: u32) -> bool {
    root(message_type).is_some()
}

/// Report every reference held by a payload of `message_type`, in field order.
pub(in crate::package) fn visit(
    message_type: u32,
    payload: &[u8],
    limits: WireLimits,
    visitor: &mut impl FnMut(RefKind, u64),
) -> Result<(), RefError> {
    let index = root(message_type).ok_or(RefError::UnsupportedType(message_type))?;
    visit_message(index, payload, limits, visitor, 0)
}

/// Collect the object references of a payload of `message_type`.
pub(in crate::package) fn object_references(
    message_type: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<Vec<u64>, RefError> {
    let mut found = Vec::new();
    visit(message_type, payload, limits, &mut |kind, identifier| {
        if kind == RefKind::Object {
            found.push(identifier);
        }
    })?;
    Ok(found)
}

/// Re-encode a payload of `message_type`, replacing each reference for which
/// `map` returns a new identifier. Returns `None` when nothing changed.
pub(in crate::package) fn rewrite(
    message_type: u32,
    payload: &[u8],
    limits: WireLimits,
    map: &mut impl FnMut(RefKind, u64) -> Option<u64>,
) -> Result<Option<Vec<u8>>, RefError> {
    let index = root(message_type).ok_or(RefError::UnsupportedType(message_type))?;
    rewrite_message(index, payload, limits, map, 0)
}

fn schema_field(index: u16, number: u32) -> Option<Field> {
    let (_name, fields) = MESSAGES[usize::from(index)];
    fields
        .binary_search_by_key(&number, |field| field.number)
        .ok()
        .map(|position| fields[position])
}

fn visit_message(
    index: u16,
    payload: &[u8],
    limits: WireLimits,
    visitor: &mut impl FnMut(RefKind, u64),
    depth: usize,
) -> Result<(), RefError> {
    if depth > MAX_DEPTH {
        return Err(RefError::Depth);
    }
    for field in parse_wire_fields(payload, limits)? {
        let Some(schema) = schema_field(index, field.number()) else {
            continue;
        };
        let inner = length_delimited(field, payload)?;
        match schema.slot {
            Slot::Reference => visitor(RefKind::Object, reference_identifier(inner, limits)?),
            Slot::DataReference => visitor(RefKind::Data, reference_identifier(inner, limits)?),
            Slot::Message(child) => visit_message(child, inner, limits, visitor, depth + 1)?,
        }
    }
    Ok(())
}

fn rewrite_message(
    index: u16,
    payload: &[u8],
    limits: WireLimits,
    map: &mut impl FnMut(RefKind, u64) -> Option<u64>,
    depth: usize,
) -> Result<Option<Vec<u8>>, RefError> {
    if depth > MAX_DEPTH {
        return Err(RefError::Depth);
    }
    let fields = parse_wire_fields(payload, limits)?;
    let mut output: Option<Vec<u8>> = None;
    for (position, field) in fields.iter().enumerate() {
        let replacement = match schema_field(index, field.number()) {
            None => None,
            Some(schema) => {
                let inner = length_delimited(*field, payload)?;
                match schema.slot {
                    Slot::Reference => rewrite_reference(inner, limits, RefKind::Object, map)?,
                    Slot::DataReference => rewrite_reference(inner, limits, RefKind::Data, map)?,
                    Slot::Message(child) => rewrite_message(child, inner, limits, map, depth + 1)?,
                }
            },
        };
        if let Some(inner) = replacement {
            if output.is_none() {
                // First change: copy every earlier field verbatim.
                let mut copied = Vec::with_capacity(payload.len() + 8);
                for earlier in &fields[..position] {
                    copied.extend_from_slice(earlier.raw(payload)?);
                }
                output = Some(copied);
            }
            if let Some(buffer) = output.as_mut() {
                buffer.extend_from_slice(field.key(payload)?);
                encode_varint_into(buffer, inner.len() as u64);
                buffer.extend_from_slice(&inner);
            }
        } else if let Some(buffer) = output.as_mut() {
            buffer.extend_from_slice(field.raw(payload)?);
        }
    }
    Ok(output)
}

fn length_delimited(field: WireField, payload: &[u8]) -> Result<&[u8], RefError> {
    if field.wire_type() != WIRE_LENGTH_DELIMITED {
        return Err(RefError::Malformed);
    }
    Ok(field.payload(payload)?)
}

/// The identifier (field 1) of a `TSP.Reference` / `TSP.DataReference`.
fn reference_identifier(payload: &[u8], limits: WireLimits) -> Result<u64, RefError> {
    let mut identifier = None;
    for field in parse_wire_fields(payload, limits)? {
        if field.number() == 1 {
            if field.wire_type() != WIRE_VARINT || identifier.is_some() {
                return Err(RefError::Malformed);
            }
            let (value, used) = decode_varint_from_bytes(field.payload(payload)?)
                .map_err(|_| RefError::Malformed)?;
            if used != field.end() - field.payload_start() {
                return Err(RefError::Malformed);
            }
            identifier = Some(value);
        }
    }
    identifier.ok_or(RefError::Malformed)
}

fn rewrite_reference(
    payload: &[u8],
    limits: WireLimits,
    kind: RefKind,
    map: &mut impl FnMut(RefKind, u64) -> Option<u64>,
) -> Result<Option<Vec<u8>>, RefError> {
    let identifier = reference_identifier(payload, limits)?;
    let Some(replacement) = map(kind, identifier) else {
        return Ok(None);
    };
    if replacement == identifier {
        return Ok(None);
    }
    let mut output = Vec::with_capacity(payload.len() + 4);
    for field in parse_wire_fields(payload, limits)? {
        if field.number() == 1 {
            output.push(0x08);
            encode_varint_into(&mut output, replacement);
        } else {
            output.extend_from_slice(field.raw(payload)?);
        }
    }
    Ok(Some(output))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn reference(identifier: u64) -> Vec<u8> {
        let mut out = vec![0x08];
        encode_varint_into(&mut out, identifier);
        out
    }

    fn field(number: u32, inner: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        encode_varint_into(&mut out, u64::from(number << 3 | 2));
        encode_varint_into(&mut out, inner.len() as u64);
        out.extend_from_slice(inner);
        out
    }

    #[test]
    fn note_reference_is_visited_and_rewritten() {
        // KN.NoteArchive { contained_storage = 1 }, plus an unknown varint field.
        let mut payload = field(1, &reference(7));
        payload.extend_from_slice(&[0x10, 0x05]);
        let limits = WireLimits::default();
        assert_eq!(object_references(15, &payload, limits), Ok(vec![7]));
        let rewritten = rewrite(15, &payload, limits, &mut |_, id| (id == 7).then_some(300))
            .unwrap()
            .unwrap();
        let mut expected = field(1, &reference(300));
        expected.extend_from_slice(&[0x10, 0x05]);
        assert_eq!(rewritten, expected);
        assert_eq!(
            rewrite(15, &payload, limits, &mut |_, _| None),
            Ok(None),
            "an unchanged payload is not re-encoded"
        );
    }

    #[test]
    fn corpus_header_references_are_found_and_rewrites_round_trip() {
        let limits = WireLimits::default();
        let mut walked_messages = 0usize;
        for (path, package) in super::super::test_support::corpus() {
            for component in package.state.source.components().iter() {
                for object in &component.archive().objects {
                    for (message, info) in object
                        .messages
                        .iter()
                        .zip(&object.archive_info.message_infos)
                    {
                        if !admitted(message.type_) {
                            continue;
                        }
                        walked_messages += 1;
                        let walked = object_references(message.type_, &message.data, limits)
                            .unwrap_or_else(|error| panic!("{path:?}: {error:?}"));
                        for listed in &info.object_references {
                            assert!(walked.contains(listed), "{path:?}: {}", message.type_);
                        }
                        assert_eq!(
                            rewrite(message.type_, &message.data, limits, &mut |_, _| None),
                            Ok(None)
                        );
                        let shifted =
                            rewrite(message.type_, &message.data, limits, &mut |_, id| {
                                Some(id + (1 << 40))
                            })
                            .expect("rewrites");
                        let Some(shifted) = shifted else {
                            assert!(walked.is_empty() || walked.iter().all(|id| *id == 0));
                            continue;
                        };
                        let restored = rewrite(message.type_, &shifted, limits, &mut |_, id| {
                            Some(id - (1 << 40))
                        })
                        .expect("rewrites back")
                        .expect("changed back");
                        assert_eq!(
                            object_references(message.type_, &restored, limits),
                            Ok(walked.clone())
                        );
                    }
                }
            }
        }
        assert!(walked_messages > 1_000);
    }

    #[test]
    fn unadmitted_types_and_malformed_references_fail() {
        let limits = WireLimits::default();
        assert_eq!(
            object_references(99_999, &[], limits),
            Err(RefError::UnsupportedType(99_999))
        );
        // A reference field encoded as a varint is malformed.
        assert_eq!(
            object_references(15, &[0x08, 0x01], limits),
            Err(RefError::Malformed)
        );
    }
}
