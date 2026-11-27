//! Small raw protobuf helpers for structural edits: an ordered field list
//! that re-encodes untouched fields byte-for-byte, and encoders for the few
//! scalar and message shapes the transactions author.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_common::wire::parse_wire_fields;

use super::StructureError;

pub(in crate::package) const VARINT: u8 = 0;
pub(in crate::package) const LENGTH_DELIMITED: u8 = 2;
pub(in crate::package) const FIXED32: u8 = 5;
pub(in crate::package) const FIXED64: u8 = 1;

/// One parsed field: number, wire type, raw encoding and payload span.
#[derive(Debug, Clone)]
pub(in crate::package) struct Field {
    pub(in crate::package) number: u32,
    pub(in crate::package) wire_type: u8,
    raw: Vec<u8>,
    payload_start: usize,
}

impl Field {
    /// The field's payload (for a length-delimited field, without the length).
    pub(in crate::package) fn payload(&self) -> &[u8] {
        &self.raw[self.payload_start..]
    }

    /// The varint value of a varint field.
    pub(in crate::package) fn varint(&self) -> Result<u64, StructureError> {
        if self.wire_type != VARINT {
            return Err(StructureError::Unexpected("varint field"));
        }
        let (value, used) =
            decode_varint_from_bytes(self.payload()).map_err(|_| StructureError::Malformed)?;
        if used != self.payload().len() {
            return Err(StructureError::Malformed);
        }
        Ok(value)
    }

    /// The value of a `float` (fixed32) field.
    pub(in crate::package) fn float(&self) -> Result<f32, StructureError> {
        let bytes: [u8; 4] = self
            .payload()
            .try_into()
            .map_err(|_| StructureError::Unexpected("float field"))?;
        if self.wire_type != FIXED32 {
            return Err(StructureError::Unexpected("float field"));
        }
        Ok(f32::from_le_bytes(bytes))
    }

    /// The payload of a length-delimited field.
    pub(in crate::package) fn message(&self) -> Result<&[u8], StructureError> {
        if self.wire_type != LENGTH_DELIMITED {
            return Err(StructureError::Unexpected("message field"));
        }
        Ok(self.payload())
    }
}

/// An ordered, editable field list.
#[derive(Debug, Clone)]
pub(in crate::package) struct Fields {
    fields: Vec<Field>,
}

impl Fields {
    /// Parse a payload.
    pub(in crate::package) fn parse(
        payload: &[u8],
        limits: WireLimits,
    ) -> Result<Self, StructureError> {
        let parsed = parse_wire_fields(payload, limits)?;
        let mut fields = Vec::with_capacity(parsed.len());
        for field in parsed {
            fields.push(Field {
                number: field.number(),
                wire_type: field.wire_type(),
                raw: field.raw(payload)?.to_vec(),
                payload_start: field.payload_start() - field.start(),
            });
        }
        Ok(Self { fields })
    }

    /// Every occurrence of field `number`.
    pub(in crate::package) fn all(&self, number: u32) -> impl Iterator<Item = &Field> {
        self.fields
            .iter()
            .filter(move |field| field.number == number)
    }

    /// The single occurrence of field `number`, if any; several are an error.
    pub(in crate::package) fn unique(&self, number: u32) -> Result<Option<&Field>, StructureError> {
        let mut found = self.all(number);
        let first = found.next();
        if found.next().is_some() {
            return Err(StructureError::Unexpected("duplicate singular field"));
        }
        Ok(first)
    }

    /// Remove every occurrence of field `number`; returns how many.
    pub(in crate::package) fn remove(&mut self, number: u32) -> usize {
        let before = self.fields.len();
        self.fields.retain(|field| field.number != number);
        before - self.fields.len()
    }

    /// Replace every occurrence of `number` with `values` (each an encoded
    /// payload of `wire_type`), placed where the first occurrence was, or
    /// appended when there was none.
    pub(in crate::package) fn replace_all(
        &mut self,
        number: u32,
        wire_type: u8,
        values: &[Vec<u8>],
    ) {
        let position = self.fields.iter().position(|field| field.number == number);
        self.fields.retain(|field| field.number != number);
        let insert_at = position.unwrap_or(self.fields.len()).min(self.fields.len());
        let encoded = values
            .iter()
            .map(|value| field(number, wire_type, value))
            .collect::<Vec<_>>();
        for (offset, item) in encoded.into_iter().enumerate() {
            self.fields.insert(insert_at + offset, item);
        }
    }

    /// Set a singular varint field, in place when present.
    pub(in crate::package) fn set_varint(&mut self, number: u32, value: u64) {
        let mut payload = Vec::new();
        encode_varint_into(&mut payload, value);
        self.set_single(number, VARINT, &payload);
    }

    /// Set a singular `float` field, in place when present.
    pub(in crate::package) fn set_float(&mut self, number: u32, value: f32) {
        self.set_single(number, FIXED32, &value.to_le_bytes());
    }

    /// Set a singular `double` field, in place when present.
    pub(in crate::package) fn set_double(&mut self, number: u32, value: f64) {
        self.set_single(number, FIXED64, &value.to_le_bytes());
    }

    /// Set a singular message field, in place when present.
    pub(in crate::package) fn set_message(&mut self, number: u32, payload: &[u8]) {
        self.set_single(number, LENGTH_DELIMITED, payload);
    }

    fn set_single(&mut self, number: u32, wire_type: u8, payload: &[u8]) {
        self.replace_all(number, wire_type, &[payload.to_vec()]);
    }

    /// Encode the field list.
    pub(in crate::package) fn encode(&self) -> Vec<u8> {
        let mut output = Vec::with_capacity(self.fields.iter().map(|field| field.raw.len()).sum());
        for field in &self.fields {
            output.extend_from_slice(&field.raw);
        }
        output
    }
}

fn field(number: u32, wire_type: u8, payload: &[u8]) -> Field {
    let mut raw = Vec::with_capacity(payload.len() + 6);
    encode_varint_into(&mut raw, u64::from(number << 3 | u32::from(wire_type)));
    if wire_type == LENGTH_DELIMITED {
        encode_varint_into(&mut raw, payload.len() as u64);
    }
    let payload_start = raw.len();
    raw.extend_from_slice(payload);
    Field {
        number,
        wire_type,
        raw,
        payload_start,
    }
}

/// An encoder for freshly authored messages.
#[derive(Debug, Default, Clone)]
pub(in crate::package) struct Writer {
    bytes: Vec<u8>,
}

impl Writer {
    pub(in crate::package) const fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn key(&mut self, number: u32, wire_type: u8) {
        encode_varint_into(
            &mut self.bytes,
            u64::from(number << 3 | u32::from(wire_type)),
        );
    }

    pub(in crate::package) fn varint(mut self, number: u32, value: u64) -> Self {
        self.key(number, VARINT);
        encode_varint_into(&mut self.bytes, value);
        self
    }

    pub(in crate::package) fn float(mut self, number: u32, value: f32) -> Self {
        self.key(number, FIXED32);
        self.bytes.extend_from_slice(&value.to_le_bytes());
        self
    }

    pub(in crate::package) fn double(mut self, number: u32, value: f64) -> Self {
        self.key(number, FIXED64);
        self.bytes.extend_from_slice(&value.to_le_bytes());
        self
    }

    pub(in crate::package) fn boolean(self, number: u32, value: bool) -> Self {
        self.varint(number, u64::from(value))
    }

    pub(in crate::package) fn bytes(mut self, number: u32, value: &[u8]) -> Self {
        self.key(number, LENGTH_DELIMITED);
        encode_varint_into(&mut self.bytes, value.len() as u64);
        self.bytes.extend_from_slice(value);
        self
    }

    pub(in crate::package) fn string(self, number: u32, value: &str) -> Self {
        self.bytes(number, value.as_bytes())
    }

    pub(in crate::package) fn message(self, number: u32, value: &Writer) -> Self {
        self.bytes(number, &value.bytes)
    }

    pub(in crate::package) fn reference(self, number: u32, identifier: u64) -> Self {
        self.message(number, &Writer::new().varint(1, identifier))
    }

    pub(in crate::package) fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

/// Encode a `TSP.Reference` payload.
pub(in crate::package) fn reference(identifier: u64) -> Vec<u8> {
    Writer::new().varint(1, identifier).finish()
}

/// The identifier of a `TSP.Reference` payload.
pub(in crate::package) fn reference_identifier(
    payload: &[u8],
    limits: WireLimits,
) -> Result<u64, StructureError> {
    let fields = Fields::parse(payload, limits)?;
    fields
        .unique(1)?
        .ok_or(StructureError::Unexpected("reference identifier"))?
        .varint()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn untouched_fields_keep_their_bytes_and_order() {
        let payload = Writer::new()
            .varint(1, 5)
            .string(2, "x")
            .varint(3, 9)
            .finish();
        let limits = WireLimits::default();
        let mut fields = Fields::parse(&payload, limits).unwrap();
        assert_eq!(fields.encode(), payload);
        fields.set_varint(3, 10);
        fields.set_message(4, &reference(7));
        let edited = Fields::parse(&fields.encode(), limits).unwrap();
        assert_eq!(edited.unique(1).unwrap().unwrap().varint().unwrap(), 5);
        assert_eq!(edited.unique(3).unwrap().unwrap().varint().unwrap(), 10);
        let reference = edited.unique(4).unwrap().unwrap().message().unwrap();
        assert_eq!(reference_identifier(reference, limits).unwrap(), 7);
        let mut removed = edited.clone();
        assert_eq!(removed.remove(2), 1);
        assert!(removed.unique(2).unwrap().is_none());
    }

    #[test]
    fn repeated_fields_are_replaced_in_place_and_writers_encode_messages() {
        let limits = WireLimits::default();
        let payload = Writer::new()
            .varint(1, 1)
            .reference(2, 10)
            .reference(2, 11)
            .varint(3, 3)
            .finish();
        let mut fields = Fields::parse(&payload, limits).unwrap();
        fields.replace_all(
            2,
            LENGTH_DELIMITED,
            &[reference(20), reference(21), reference(22)],
        );
        let numbers = Fields::parse(&fields.encode(), limits)
            .unwrap()
            .fields
            .iter()
            .map(|field| field.number)
            .collect::<Vec<_>>();
        assert_eq!(numbers, [1, 2, 2, 2, 3]);
        let identifiers = Fields::parse(&fields.encode(), limits)
            .unwrap()
            .all(2)
            .map(|field| reference_identifier(field.message().unwrap(), limits).unwrap())
            .collect::<Vec<_>>();
        assert_eq!(identifiers, [20, 21, 22]);
        fields.replace_all(9, LENGTH_DELIMITED, &[b"x".to_vec()]);
        assert_eq!(
            Fields::parse(&fields.encode(), limits)
                .unwrap()
                .fields
                .last()
                .unwrap()
                .number,
            9
        );
        assert!(
            Fields::parse(&fields.encode(), limits)
                .unwrap()
                .unique(2)
                .is_err()
        );
        let nested = Writer::new()
            .message(1, &Writer::new().string(1, "").varint(2, 0))
            .finish();
        assert_eq!(nested, [0x0a, 0x04, 0x0a, 0x00, 0x10, 0x00]);
    }
}
