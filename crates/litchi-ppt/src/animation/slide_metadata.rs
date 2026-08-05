//! PowerPoint 10 slide flags and creation-time atoms (MS-PPT 2.5.30-2.5.31).

use super::types::{Flags, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::package::Result;
use crate::{Error, Record};

const HEADER_LEN: usize = 8;
const FLAGS_PAYLOAD_LEN: usize = 4;
const TIME_PAYLOAD_LEN: usize = 8;
const DEFINED_FLAGS_MASK: u32 = 0x0000_0003;

/// Resource bounds for parsing PowerPoint 10 slide metadata atoms.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMetadataLimits {
    /// Maximum accepted size of one complete record, including its header.
    pub max_record_bytes: usize,
}

impl Default for SlideMetadataLimits {
    fn default() -> Self {
        Self {
            max_record_bytes: HEADER_LEN + TIME_PAYLOAD_LEN,
        }
    }
}

impl Flags {
    /// Creates flags with all undefined bits cleared.
    pub fn new(preserve_master: bool, override_master_animation: bool) -> Self {
        let raw = u32::from(preserve_master) | (u32::from(override_master_animation) << 1);
        Self {
            raw,
            preserve_master,
            override_master_animation,
        }
    }

    /// Decodes the two defined flags while retaining the ignored bits losslessly.
    pub fn from_raw(raw: u32) -> Self {
        Self {
            raw,
            preserve_master: raw & 1 != 0,
            override_master_animation: raw & 2 != 0,
        }
    }

    /// Returns the serialized word. Undefined bits are preserved; defined bits
    /// are normalized from the corresponding semantic fields.
    pub fn raw_value(&self) -> u32 {
        (self.raw & !DEFINED_FLAGS_MASK)
            | u32::from(self.preserve_master)
            | (u32::from(self.override_master_animation) << 1)
    }

    /// Returns bits 2-31, which MS-PPT requires readers to ignore.
    pub fn ignored_bits(&self) -> u32 {
        self.raw & !DEFINED_FLAGS_MASK
    }

    /// Parses a generic record using the default resource bounds.
    pub fn parse_record(record: &Record) -> Result<Self> {
        Self::parse_record_with_limits(record, SlideMetadataLimits::default())
    }

    /// Parses a generic record using explicit resource bounds.
    pub fn parse_record_with_limits(record: &Record, limits: SlideMetadataLimits) -> Result<Self> {
        let payload = validate_record(
            record,
            RecordType::SlideFlags10Atom,
            FLAGS_PAYLOAD_LEN,
            limits,
            "SlideFlags10Atom",
        )?;
        Ok(Self::from_raw(u32::from_le_bytes(
            payload.try_into().map_err(|_| {
                Error::Corrupted("SlideFlags10Atom payload is truncated".to_string())
            })?,
        )))
    }

    /// Parses exactly one serialized record using the default resource bounds.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, SlideMetadataLimits::default())
    }

    /// Parses exactly one serialized record using explicit resource bounds.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: SlideMetadataLimits) -> Result<Self> {
        let payload = validate_bytes(
            bytes,
            RecordType::SlideFlags10Atom,
            FLAGS_PAYLOAD_LEN,
            limits,
            "SlideFlags10Atom",
        )?;
        Ok(Self::from_raw(u32::from_le_bytes(
            payload.try_into().map_err(|_| {
                Error::Corrupted("SlideFlags10Atom payload is truncated".to_string())
            })?,
        )))
    }

    /// Serializes the fixed four-byte payload.
    pub fn to_payload(&self) -> [u8; FLAGS_PAYLOAD_LEN] {
        self.raw_value().to_le_bytes()
    }

    /// Serializes the complete atom, including its record header.
    pub fn to_bytes(&self) -> Vec<u8> {
        serialize_atom(RecordType::SlideFlags10Atom, &self.to_payload())
    }

    /// Converts the atom into the generic record representation.
    pub fn to_record(&self) -> Record {
        let data = self.to_payload().to_vec();
        Record {
            record_type: RecordType::SlideFlags10Atom,
            record_type_raw: RecordType::SlideFlags10Atom.as_u16(),
            version: 0,
            instance: 0,
            data_length: FLAGS_PAYLOAD_LEN as u32,
            data,
            children: Vec::new(),
        }
    }
}

/// A `SlideTime10Atom` creation timestamp expressed as a Windows FILETIME.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTime {
    file_time: u64,
}

impl SlideTime {
    /// Creates an atom from 100-nanosecond ticks since 1601-01-01 UTC.
    pub const fn new(file_time: u64) -> Self {
        Self { file_time }
    }

    /// Returns the raw FILETIME value.
    pub const fn file_time(self) -> u64 {
        self.file_time
    }

    /// Parses a generic record using the default resource bounds.
    pub fn parse_record(record: &Record) -> Result<Self> {
        Self::parse_record_with_limits(record, SlideMetadataLimits::default())
    }

    /// Parses a generic record using explicit resource bounds.
    pub fn parse_record_with_limits(record: &Record, limits: SlideMetadataLimits) -> Result<Self> {
        let payload = validate_record(
            record,
            RecordType::SlideTime10Atom,
            TIME_PAYLOAD_LEN,
            limits,
            "SlideTime10Atom",
        )?;
        Ok(Self::new(u64::from_le_bytes(payload.try_into().map_err(
            |_| Error::Corrupted("SlideTime10Atom payload is truncated".to_string()),
        )?)))
    }

    /// Parses exactly one serialized record using the default resource bounds.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, SlideMetadataLimits::default())
    }

    /// Parses exactly one serialized record using explicit resource bounds.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: SlideMetadataLimits) -> Result<Self> {
        let payload = validate_bytes(
            bytes,
            RecordType::SlideTime10Atom,
            TIME_PAYLOAD_LEN,
            limits,
            "SlideTime10Atom",
        )?;
        Ok(Self::new(u64::from_le_bytes(payload.try_into().map_err(
            |_| Error::Corrupted("SlideTime10Atom payload is truncated".to_string()),
        )?)))
    }

    /// Serializes the fixed eight-byte FILETIME payload.
    pub fn to_payload(self) -> [u8; TIME_PAYLOAD_LEN] {
        self.file_time.to_le_bytes()
    }

    /// Serializes the complete atom, including its record header.
    pub fn to_bytes(self) -> Vec<u8> {
        serialize_atom(RecordType::SlideTime10Atom, &self.to_payload())
    }

    /// Converts the atom into the generic record representation.
    pub fn to_record(self) -> Record {
        let data = self.to_payload().to_vec();
        Record {
            record_type: RecordType::SlideTime10Atom,
            record_type_raw: RecordType::SlideTime10Atom.as_u16(),
            version: 0,
            instance: 0,
            data_length: TIME_PAYLOAD_LEN as u32,
            data,
            children: Vec::new(),
        }
    }
}

impl SlideAnimationExtension {
    /// Returns the typed `SlideFlags10Atom`, if present.
    pub fn slide_flags_atom(&self) -> Option<Flags> {
        self.slide_flags
    }

    /// Replaces the optional typed `SlideFlags10Atom`.
    pub fn set_slide_flags_atom(&mut self, flags: Flags) {
        self.slide_flags = Some(flags);
    }

    /// Returns the typed `SlideTime10Atom`, if present.
    pub fn slide_time_atom(&self) -> Option<SlideTime> {
        self.creation_time_filetime.map(SlideTime::new)
    }

    /// Replaces the optional typed `SlideTime10Atom`.
    pub fn set_slide_time_atom(&mut self, time: SlideTime) {
        self.creation_time_filetime = Some(time.file_time());
    }
}

fn validate_record<'a>(
    record: &'a Record,
    expected_type: RecordType,
    payload_len: usize,
    limits: SlideMetadataLimits,
    name: &str,
) -> Result<&'a [u8]> {
    let total_len = HEADER_LEN + payload_len;
    if total_len > limits.max_record_bytes {
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
    if record.data_length != payload_len as u32 || record.data.len() != payload_len {
        return Err(Error::InvalidFormat(format!(
            "{name} requires a {payload_len}-byte payload"
        )));
    }
    if !record.children.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{name} is an atom and cannot contain child records"
        )));
    }
    Ok(&record.data)
}

fn validate_bytes<'a>(
    bytes: &'a [u8],
    expected_type: RecordType,
    payload_len: usize,
    limits: SlideMetadataLimits,
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
    let version = version_instance & 0x000f;
    let instance = version_instance >> 4;
    let record_type = u16::from_le_bytes([bytes[2], bytes[3]]);
    let declared_len = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);

    if version != 0 || instance != 0 {
        return Err(Error::InvalidFormat(format!(
            "{name} requires record version 0 and instance 0"
        )));
    }
    if record_type != expected_type.as_u16() {
        return Err(Error::InvalidFormat(format!("expected {name} record type")));
    }
    if declared_len != payload_len as u32 {
        return Err(Error::InvalidFormat(format!(
            "{name} requires a {payload_len}-byte payload"
        )));
    }

    let expected_len = HEADER_LEN + payload_len;
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

fn serialize_atom(record_type: RecordType, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(HEADER_LEN + payload.len());
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&record_type.as_u16().to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn slide_flags_exhaustive_defined_bits_and_ignored_bits_round_trip() {
        for defined in 0..=3u32 {
            for ignored in [0, 0x5555_5554, 0xaaaa_aaa8, 0xffff_fffc] {
                let raw = defined | ignored;
                let flags = Flags::from_raw(raw);
                assert_eq!(flags.preserve_master, defined & 1 != 0);
                assert_eq!(flags.override_master_animation, defined & 2 != 0);
                assert_eq!(flags.ignored_bits(), ignored);
                assert_eq!(flags.raw_value(), raw);
                assert_eq!(Flags::parse_bytes(&flags.to_bytes()).unwrap(), flags);
                assert_eq!(Flags::parse_record(&flags.to_record()).unwrap(), flags);
            }
        }
    }

    #[test]
    fn slide_flags_semantic_fields_normalize_only_defined_bits() {
        let flags = Flags {
            raw: 0xffff_ffff,
            preserve_master: false,
            override_master_animation: true,
        };
        assert_eq!(flags.raw_value(), 0xffff_fffe);
        assert_eq!(&flags.to_bytes()[8..], &0xffff_fffeu32.to_le_bytes());
    }

    #[test]
    fn slide_time_all_bits_round_trip_and_extension_accessors() {
        for file_time in [0, 1, 0x0123_4567_89ab_cdef, u64::MAX] {
            let time = SlideTime::new(file_time);
            assert_eq!(SlideTime::parse_bytes(&time.to_bytes()).unwrap(), time);
            assert_eq!(SlideTime::parse_record(&time.to_record()).unwrap(), time);
        }

        let mut extension = SlideAnimationExtension::default();
        let flags = Flags::from_raw(0xffff_fffd);
        let time = SlideTime::new(u64::MAX);
        extension.set_slide_flags_atom(flags);
        extension.set_slide_time_atom(time);
        assert_eq!(extension.slide_flags_atom(), Some(flags));
        assert_eq!(extension.slide_time_atom(), Some(time));
    }

    #[test]
    fn both_atoms_reject_every_header_field_violation() {
        let cases = [
            (Flags::new(false, false).to_bytes(), 4usize),
            (SlideTime::new(0).to_bytes(), 8usize),
        ];
        for (valid, payload_len) in cases {
            let parse = |bytes: &[u8]| {
                if payload_len == 4 {
                    Flags::parse_bytes(bytes).map(|_| ())
                } else {
                    SlideTime::parse_bytes(bytes).map(|_| ())
                }
            };

            let mut bad = valid.clone();
            bad[0] = 1;
            assert!(parse(&bad).is_err());
            let mut bad = valid.clone();
            bad[1] = 0x10;
            assert!(parse(&bad).is_err());
            let mut bad = valid.clone();
            bad[2..4].copy_from_slice(&0xffffu16.to_le_bytes());
            assert!(parse(&bad).is_err());
            let mut bad = valid.clone();
            bad[4..8].copy_from_slice(&((payload_len - 1) as u32).to_le_bytes());
            assert!(parse(&bad).is_err());
            let mut bad = valid.clone();
            bad[4..8].copy_from_slice(&((payload_len + 1) as u32).to_le_bytes());
            assert!(parse(&bad).is_err());
        }
    }

    #[test]
    fn both_atoms_reject_all_truncations_trailing_data_and_resource_overruns() {
        let flags = Flags::new(true, true).to_bytes();
        for end in 0..flags.len() {
            assert!(Flags::parse_bytes(&flags[..end]).is_err());
        }
        let mut trailing = flags.clone();
        trailing.push(0);
        assert!(Flags::parse_bytes(&trailing).is_err());
        assert!(
            Flags::parse_bytes_with_limits(
                &flags,
                SlideMetadataLimits {
                    max_record_bytes: flags.len() - 1,
                },
            )
            .is_err()
        );

        let time = SlideTime::new(u64::MAX).to_bytes();
        for end in 0..time.len() {
            assert!(SlideTime::parse_bytes(&time[..end]).is_err());
        }
        let mut trailing = time.clone();
        trailing.push(0);
        assert!(SlideTime::parse_bytes(&trailing).is_err());
        assert!(
            SlideTime::parse_bytes_with_limits(
                &time,
                SlideMetadataLimits {
                    max_record_bytes: time.len() - 1,
                },
            )
            .is_err()
        );
    }

    #[test]
    fn generic_records_reject_inconsistent_type_length_and_children() {
        let mut flags = Flags::new(false, false).to_record();
        flags.record_type_raw = RecordType::SlideTime10Atom.as_u16();
        assert!(Flags::parse_record(&flags).is_err());

        let mut time = SlideTime::new(0).to_record();
        time.data_length = 7;
        assert!(SlideTime::parse_record(&time).is_err());

        let mut time = SlideTime::new(0).to_record();
        time.children.push(Flags::new(false, false).to_record());
        assert!(SlideTime::parse_record(&time).is_err());

        let mut flags = Flags::new(false, false).to_record();
        flags.record_type = RecordType::SlideTime10Atom;
        assert!(Flags::parse_record(&flags).is_err());
    }
}
