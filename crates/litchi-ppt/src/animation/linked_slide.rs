//! Inert `PowerPoint` 10 linked-slide records (MS-PPT 2.5.32-2.5.33).

use super::types::SlideAnimationExtension;
use crate::consts::RecordType;
use crate::package::Result;
use crate::{Error, Record};

const HEADER_LEN: usize = 8;
const PAYLOAD_LEN: usize = 8;

/// Resource bounds for linked-slide metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LinkedSlideLimits {
    /// Maximum accepted size of one complete atom, including its header.
    pub max_record_bytes: usize,
    /// Maximum accepted `LinkedShape10Atom` count.
    pub max_linked_shapes: usize,
}

impl Default for LinkedSlideLimits {
    fn default() -> Self {
        Self {
            max_record_bytes: HEADER_LEN + PAYLOAD_LEN,
            max_linked_shapes: 65_536,
        }
    }
}

/// A `LinkedSlide10Atom` containing an inert cross-document slide identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LinkedSlide {
    linked_slide_id_ref: u32,
    linked_shape_count: u32,
}

impl LinkedSlide {
    /// Creates an atom without resolving or opening the referenced document.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(linked_slide_id_ref: u32, linked_shape_count: u32) -> Result<Self> {
        if linked_shape_count > i32::MAX as u32 {
            return Err(Error::InvalidFormat(
                "LinkedSlide10Atom shape count exceeds signed 32-bit range".to_string(),
            ));
        }
        Ok(Self {
            linked_slide_id_ref,
            linked_shape_count,
        })
    }

    /// Returns the slide identifier; zero is the normative null reference.
    #[must_use]
    pub const fn linked_slide_id_ref(self) -> u32 {
        self.linked_slide_id_ref
    }

    /// Returns the declared number of immediately following shape atoms.
    #[must_use]
    pub const fn linked_shape_count(self) -> u32 {
        self.linked_shape_count
    }

    /// Parses a generic record using the default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the record type, header fields, or payload do not
    /// match the fixed `LinkedSlide10Atom` layout, or if the shape count is
    /// negative or exceeds the configured limits.
    pub fn parse_record(record: &Record) -> Result<Self> {
        Self::parse_record_with_limits(record, LinkedSlideLimits::default())
    }

    /// Parses a generic record using explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the record type, header fields, or payload do not
    /// match the fixed `LinkedSlide10Atom` layout, or if the shape count is
    /// negative or exceeds the given limits.
    pub fn parse_record_with_limits(record: &Record, limits: LinkedSlideLimits) -> Result<Self> {
        let payload = validate_record(
            record,
            RecordType::LinkedSlide10Atom,
            limits,
            "LinkedSlide10Atom",
        )?;
        Self::parse_payload(payload, limits)
    }

    /// Parses exactly one serialized record using the default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not exactly one `LinkedSlide10Atom`
    /// record, or if the shape count is negative or exceeds the configured
    /// limits.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, LinkedSlideLimits::default())
    }

    /// Parses exactly one serialized record using explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not exactly one `LinkedSlide10Atom`
    /// record, or if the shape count is negative or exceeds the given limits.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: LinkedSlideLimits) -> Result<Self> {
        let payload = validate_bytes(
            bytes,
            RecordType::LinkedSlide10Atom,
            limits,
            "LinkedSlide10Atom",
        )?;
        Self::parse_payload(payload, limits)
    }

    fn parse_payload(payload: &[u8], limits: LinkedSlideLimits) -> Result<Self> {
        let linked_slide_id_ref = u32::from_le_bytes(payload[0..4].try_into().map_err(|_err| {
            Error::Corrupted("LinkedSlide10Atom slide identifier is truncated".to_string())
        })?);
        let signed_count = i32::from_le_bytes(payload[4..8].try_into().map_err(|_err| {
            Error::Corrupted("LinkedSlide10Atom shape count is truncated".to_string())
        })?);
        let linked_shape_count = u32::try_from(signed_count).map_err(|_err| {
            Error::InvalidFormat("LinkedSlide10Atom shape count cannot be negative".to_string())
        })?;
        let count = usize::try_from(linked_shape_count).map_err(|_err| {
            Error::InvalidFormat(
                "LinkedSlide10Atom shape count does not fit this platform".to_string(),
            )
        })?;
        if count > limits.max_linked_shapes {
            return Err(Error::InvalidFormat(format!(
                "LinkedSlide10Atom shape count {count} exceeds configured limit {}",
                limits.max_linked_shapes
            )));
        }
        Self::new(linked_slide_id_ref, linked_shape_count)
    }

    #[must_use]
    pub fn to_payload(self) -> [u8; PAYLOAD_LEN] {
        let mut payload = [0; PAYLOAD_LEN];
        payload[..4].copy_from_slice(&self.linked_slide_id_ref.to_le_bytes());
        payload[4..].copy_from_slice(&self.linked_shape_count.to_le_bytes());
        payload
    }

    #[must_use]
    pub fn to_bytes(self) -> Vec<u8> {
        serialize_atom(RecordType::LinkedSlide10Atom, &self.to_payload())
    }

    #[must_use]
    pub fn to_record(self) -> Record {
        generic_record(RecordType::LinkedSlide10Atom, self.to_payload().to_vec())
    }
}

/// A `LinkedShape10Atom` containing two inert shape identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LinkedShape {
    shape_id_ref: u32,
    linked_shape_id_ref: u32,
}

impl LinkedShape {
    #[must_use]
    pub const fn new(shape_id_ref: u32, linked_shape_id_ref: u32) -> Self {
        Self {
            shape_id_ref,
            linked_shape_id_ref,
        }
    }

    #[must_use]
    pub const fn shape_id_ref(self) -> u32 {
        self.shape_id_ref
    }

    #[must_use]
    pub const fn linked_shape_id_ref(self) -> u32 {
        self.linked_shape_id_ref
    }

    /// Parses a generic record using the default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the record type, header fields, or payload do not
    /// match the fixed `LinkedShape10Atom` layout.
    pub fn parse_record(record: &Record) -> Result<Self> {
        Self::parse_record_with_limits(record, LinkedSlideLimits::default())
    }

    /// Parses a generic record using explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the record type, header fields, or payload do not
    /// match the fixed `LinkedShape10Atom` layout.
    pub fn parse_record_with_limits(record: &Record, limits: LinkedSlideLimits) -> Result<Self> {
        Self::parse_payload(validate_record(
            record,
            RecordType::LinkedShape10Atom,
            limits,
            "LinkedShape10Atom",
        )?)
    }

    /// Parses exactly one serialized record using the default resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not exactly one `LinkedShape10Atom`
    /// record.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, LinkedSlideLimits::default())
    }

    /// Parses exactly one serialized record using explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not exactly one `LinkedShape10Atom`
    /// record.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: LinkedSlideLimits) -> Result<Self> {
        Self::parse_payload(validate_bytes(
            bytes,
            RecordType::LinkedShape10Atom,
            limits,
            "LinkedShape10Atom",
        )?)
    }

    fn parse_payload(payload: &[u8]) -> Result<Self> {
        let shape_id_ref = u32::from_le_bytes(payload[0..4].try_into().map_err(|_err| {
            Error::Corrupted("LinkedShape10Atom shape identifier is truncated".to_string())
        })?);
        let linked_shape_id_ref = u32::from_le_bytes(payload[4..8].try_into().map_err(|_err| {
            Error::Corrupted("LinkedShape10Atom linked-shape identifier is truncated".to_string())
        })?);
        Ok(Self::new(shape_id_ref, linked_shape_id_ref))
    }

    #[must_use]
    pub fn to_payload(self) -> [u8; PAYLOAD_LEN] {
        let mut payload = [0; PAYLOAD_LEN];
        payload[..4].copy_from_slice(&self.shape_id_ref.to_le_bytes());
        payload[4..].copy_from_slice(&self.linked_shape_id_ref.to_le_bytes());
        payload
    }

    #[must_use]
    pub fn to_bytes(self) -> Vec<u8> {
        serialize_atom(RecordType::LinkedShape10Atom, &self.to_payload())
    }

    #[must_use]
    pub fn to_record(self) -> Record {
        generic_record(RecordType::LinkedShape10Atom, self.to_payload().to_vec())
    }
}

impl SlideAnimationExtension {
    #[must_use]
    pub fn linked_slide_atom(&self) -> Option<LinkedSlide> {
        self.linked_slide
    }

    pub fn set_linked_slide_atom(&mut self, linked_slide: Option<LinkedSlide>) {
        self.linked_slide = linked_slide;
    }

    #[must_use]
    pub fn linked_shape_atoms(&self) -> &[LinkedShape] {
        &self.linked_shapes
    }

    pub fn set_linked_shape_atoms(&mut self, linked_shapes: Vec<LinkedShape>) {
        self.linked_shapes = linked_shapes;
    }
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "`PAYLOAD_LEN` is a compile-time 8-byte constant, so it always fits in `u32`"
)]
fn validate_record<'a>(
    record: &'a Record,
    expected_type: RecordType,
    limits: LinkedSlideLimits,
    name: &str,
) -> Result<&'a [u8]> {
    if HEADER_LEN + PAYLOAD_LEN > limits.max_record_bytes {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the configured record-size limit"
        )));
    }
    if record.record_type != expected_type || record.record_type_raw != expected_type.as_u16() {
        return Err(Error::InvalidFormat(format!("expected {name} record type")));
    }
    if record.version != 0 || record.instance != 0 {
        return Err(Error::InvalidFormat(format!(
            "{name} requires record version 0 and instance 0"
        )));
    }
    if record.data_length != PAYLOAD_LEN as u32 || record.data.len() != PAYLOAD_LEN {
        return Err(Error::InvalidFormat(format!(
            "{name} requires an eight-byte payload"
        )));
    }
    if !record.children.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{name} is an atom and cannot contain child records"
        )));
    }
    Ok(&record.data)
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "`PAYLOAD_LEN` is a compile-time 8-byte constant, so it always fits in `u32`"
)]
fn validate_bytes<'a>(
    bytes: &'a [u8],
    expected_type: RecordType,
    limits: LinkedSlideLimits,
    name: &str,
) -> Result<&'a [u8]> {
    if bytes.len() > limits.max_record_bytes {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the configured record-size limit"
        )));
    }
    if bytes.len() < HEADER_LEN {
        return Err(Error::Corrupted(format!(
            "{name} record header is truncated"
        )));
    }
    let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
    if version_instance & 0x000f != 0 || version_instance >> 4 != 0 {
        return Err(Error::InvalidFormat(format!(
            "{name} requires record version 0 and instance 0"
        )));
    }
    if u16::from_le_bytes([bytes[2], bytes[3]]) != expected_type.as_u16() {
        return Err(Error::InvalidFormat(format!("expected {name} record type")));
    }
    if u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]) != PAYLOAD_LEN as u32 {
        return Err(Error::InvalidFormat(format!(
            "{name} requires an eight-byte payload"
        )));
    }
    let expected_len = HEADER_LEN + PAYLOAD_LEN;
    if bytes.len() < expected_len {
        return Err(Error::Corrupted(format!("{name} payload is truncated")));
    }
    if bytes.len() > expected_len {
        return Err(Error::InvalidFormat(format!(
            "{name} record has trailing data"
        )));
    }
    Ok(&bytes[HEADER_LEN..expected_len])
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "callers pass only the fixed eight-byte atom payloads, so the payload length always fits in `u32`"
)]
fn serialize_atom(record_type: RecordType, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(HEADER_LEN + payload.len());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&record_type.as_u16().to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "`PAYLOAD_LEN` is a compile-time 8-byte constant, so it always fits in `u32`"
)]
fn generic_record(record_type: RecordType, data: Vec<u8>) -> Record {
    Record {
        record_type,
        record_type_raw: record_type.as_u16(),
        version: 0,
        instance: 0,
        data_length: PAYLOAD_LEN as u32,
        data,
        children: Vec::new(),
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::animation::{Flags, parse_slide_animation_extension};

    #[test]
    fn linked_slide_round_trips_and_enforces_signed_bounded_count() {
        for (id, count) in [(0, 0), (1, 1), (u32::MAX, 65_536)] {
            let atom = LinkedSlide::new(id, count).unwrap();
            assert_eq!(LinkedSlide::parse_bytes(&atom.to_bytes()).unwrap(), atom);
            assert_eq!(LinkedSlide::parse_record(&atom.to_record()).unwrap(), atom);
        }
        assert!(LinkedSlide::new(0, i32::MAX as u32 + 1).is_err());
        let mut negative = LinkedSlide::new(0, 0).unwrap().to_bytes();
        negative[12..16].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(LinkedSlide::parse_bytes(&negative).is_err());
        let over_limit = LinkedSlide::new(0, 3).unwrap().to_bytes();
        assert!(
            LinkedSlide::parse_bytes_with_limits(
                &over_limit,
                LinkedSlideLimits {
                    max_record_bytes: 16,
                    max_linked_shapes: 2,
                },
            )
            .is_err()
        );
    }

    #[test]
    fn linked_shape_round_trips_all_identifier_bits() {
        for atom in [
            LinkedShape::new(0, 0),
            LinkedShape::new(1, 2),
            LinkedShape::new(u32::MAX, 0x8000_0000),
        ] {
            assert_eq!(LinkedShape::parse_bytes(&atom.to_bytes()).unwrap(), atom);
            assert_eq!(LinkedShape::parse_record(&atom.to_record()).unwrap(), atom);
        }
    }

    #[test]
    fn both_atoms_reject_headers_truncations_trailing_and_generic_children() {
        let valid_records = [
            LinkedSlide::new(0, 0).unwrap().to_bytes(),
            LinkedShape::new(0, 0).to_bytes(),
        ];
        for (index, valid) in valid_records.into_iter().enumerate() {
            let parse = |bytes: &[u8]| {
                if index == 0 {
                    LinkedSlide::parse_bytes(bytes).map(|_| ())
                } else {
                    LinkedShape::parse_bytes(bytes).map(|_| ())
                }
            };
            for end in 0..valid.len() {
                assert!(parse(&valid[..end]).is_err());
            }
            for (offset, value) in [(0, 1), (1, 0x10)] {
                let mut bad = valid.clone();
                bad[offset] = value;
                assert!(parse(&bad).is_err());
            }
            let mut bad_type = valid.clone();
            bad_type[2..4].copy_from_slice(&0xffffu16.to_le_bytes());
            assert!(parse(&bad_type).is_err());
            let mut bad_length = valid.clone();
            bad_length[4..8].copy_from_slice(&7u32.to_le_bytes());
            assert!(parse(&bad_length).is_err());
            let mut trailing = valid;
            trailing.push(0);
            let limits = LinkedSlideLimits {
                max_record_bytes: 17,
                ..LinkedSlideLimits::default()
            };
            let result = if index == 0 {
                LinkedSlide::parse_bytes_with_limits(&trailing, limits).map(|_| ())
            } else {
                LinkedShape::parse_bytes_with_limits(&trailing, limits).map(|_| ())
            };
            assert!(result.is_err());
        }

        let mut record = LinkedShape::new(0, 0).to_record();
        record.children.push(LinkedShape::new(1, 1).to_record());
        assert!(LinkedShape::parse_record(&record).is_err());
    }

    #[test]
    fn extension_enforces_declared_count_order_and_contiguity() {
        let linked_slide = LinkedSlide::new(7, 2).unwrap();
        let shapes = [LinkedShape::new(10, 20), LinkedShape::new(11, 21)];
        let mut bytes = linked_slide.to_bytes();
        bytes.extend_from_slice(&shapes[0].to_bytes());
        bytes.extend_from_slice(&shapes[1].to_bytes());
        let parsed = parse_slide_animation_extension(&bytes).unwrap();
        assert_eq!(parsed.linked_slide_atom(), Some(linked_slide));
        assert_eq!(parsed.linked_shape_atoms(), &shapes);

        let mut missing = LinkedSlide::new(7, 2).unwrap().to_bytes();
        missing.extend_from_slice(&shapes[0].to_bytes());
        assert!(parse_slide_animation_extension(&missing).is_err());
        let mut excess = LinkedSlide::new(7, 1).unwrap().to_bytes();
        excess.extend_from_slice(&shapes[0].to_bytes());
        excess.extend_from_slice(&shapes[1].to_bytes());
        assert!(parse_slide_animation_extension(&excess).is_err());
        let mut before = shapes[0].to_bytes();
        before.extend_from_slice(&LinkedSlide::new(7, 1).unwrap().to_bytes());
        assert!(parse_slide_animation_extension(&before).is_err());
        let mut split = LinkedSlide::new(7, 2).unwrap().to_bytes();
        split.extend_from_slice(&shapes[0].to_bytes());
        split.extend_from_slice(&Flags::new(false, false).to_bytes());
        split.extend_from_slice(&shapes[1].to_bytes());
        assert!(parse_slide_animation_extension(&split).is_err());
    }
}
