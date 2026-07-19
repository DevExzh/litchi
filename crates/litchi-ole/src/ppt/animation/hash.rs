//! MS-PPT §2.8.3 PowerPoint 2002 animation hash atom.

use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;

const HEADER_LEN: usize = 8;
const PAYLOAD_LEN: usize = 4;
const RECORD_LEN: usize = HEADER_LEN + PAYLOAD_LEN;

/// Resource limit for parsing a HashCode10Atom record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointAnimationHash10Limits {
    /// Maximum accepted complete record size.
    pub max_record_bytes: usize,
}

impl Default for PowerPointAnimationHash10Limits {
    fn default() -> Self {
        Self {
            max_record_bytes: RECORD_LEN,
        }
    }
}

/// Exact MS-PPT §2.8.3 HashCode10Atom value.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointAnimationHash10 {
    hash: u32,
}

impl PowerPointAnimationHash10 {
    /// Construct an inert animation hash value.
    pub const fn new(hash: u32) -> Self {
        Self { hash }
    }

    /// Raw hash value stored by the producing application.
    pub const fn hash(self) -> u32 {
        self.hash
    }

    /// Parse from the generic PowerPoint record model.
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::HashCode10Atom
            || record.record_type_raw != PptRecordType::HashCode10Atom.as_u16()
        {
            return corrupted("HashCode10Atom has an invalid record type");
        }
        if record.version != 0 {
            return corrupted("HashCode10Atom recVer must be zero");
        }
        if record.instance != 0 {
            return corrupted("HashCode10Atom recInstance must be zero");
        }
        if record.data_length != PAYLOAD_LEN as u32 || record.data.len() != PAYLOAD_LEN {
            return corrupted("HashCode10Atom recLen must be four bytes");
        }
        if !record.children.is_empty() {
            return corrupted("HashCode10Atom must not contain child records");
        }
        Ok(Self::new(u32::from_le_bytes(
            record.data[..4].try_into().expect("length checked"),
        )))
    }

    /// Parse one exact complete record.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, PowerPointAnimationHash10Limits::default())
    }

    /// Parse one exact complete record with a caller-supplied input bound.
    pub fn parse_bytes_with_limits(
        bytes: &[u8],
        limits: PowerPointAnimationHash10Limits,
    ) -> Result<Self> {
        if bytes.len() > limits.max_record_bytes {
            return corrupted("HashCode10Atom exceeds the configured record size limit");
        }
        if bytes.len() != RECORD_LEN {
            return corrupted(format!(
                "HashCode10Atom record must be exactly {RECORD_LEN} bytes, got {}",
                bytes.len()
            ));
        }
        let (record, consumed) = PptRecord::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return corrupted("HashCode10Atom has trailing bytes");
        }
        Self::parse_record(&record)
    }

    /// Exact four-byte payload.
    pub const fn to_payload(self) -> [u8; PAYLOAD_LEN] {
        self.hash.to_le_bytes()
    }

    /// Serialize the complete normative record.
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(RECORD_LEN);
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&PptRecordType::HashCode10Atom.as_u16().to_le_bytes());
        bytes.extend_from_slice(&(PAYLOAD_LEN as u32).to_le_bytes());
        bytes.extend_from_slice(&self.to_payload());
        bytes
    }

    /// Convert to the generic PowerPoint record representation.
    pub fn to_record(self) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::HashCode10Atom,
            record_type_raw: PptRecordType::HashCode10Atom.as_u16(),
            version: 0,
            instance: 0,
            data_length: PAYLOAD_LEN as u32,
            data: self.to_payload().to_vec(),
            children: Vec::new(),
        }
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::animation::SlideAnimationExtension;

    #[test]
    fn exact_round_trip_and_extension_accessors_cover_all_hash_bits() {
        for hash in [0, 1, 0x8000_0000, u32::MAX] {
            let atom = PowerPointAnimationHash10::new(hash);
            let bytes = atom.to_bytes();
            assert_eq!(
                PowerPointAnimationHash10::parse_bytes(&bytes).unwrap(),
                atom
            );
            assert_eq!(
                PowerPointAnimationHash10::parse_record(&atom.to_record()).unwrap(),
                atom
            );

            let mut extension = SlideAnimationExtension::default();
            extension.set_animation_hash_atom(Some(atom));
            assert_eq!(extension.animation_hash, Some(hash));
            assert_eq!(extension.animation_hash_atom(), Some(atom));
            extension.set_animation_hash_atom(None);
            assert_eq!(extension.animation_hash_atom(), None);
        }
    }

    #[test]
    fn rejects_every_invalid_header_field_and_declared_length() {
        let valid = PowerPointAnimationHash10::new(0x1234_5678).to_bytes();
        for (index, replacement) in [(0usize, 1u8), (1, 1), (2, 1), (3, 1)] {
            let mut bad = valid.clone();
            bad[index] = replacement;
            assert!(PowerPointAnimationHash10::parse_bytes(&bad).is_err());
        }
        for declared in [0u32, 3, 5, u32::MAX] {
            let mut bad = valid.clone();
            bad[4..8].copy_from_slice(&declared.to_le_bytes());
            assert!(PowerPointAnimationHash10::parse_bytes(&bad).is_err());
        }
    }

    #[test]
    fn rejects_truncation_trailing_bytes_and_resource_limit() {
        let valid = PowerPointAnimationHash10::new(7).to_bytes();
        for length in 0..valid.len() {
            assert!(PowerPointAnimationHash10::parse_bytes(&valid[..length]).is_err());
        }
        let mut trailing = valid.clone();
        trailing.push(0);
        assert!(PowerPointAnimationHash10::parse_bytes(&trailing).is_err());
        assert!(
            PowerPointAnimationHash10::parse_bytes_with_limits(
                &valid,
                PowerPointAnimationHash10Limits {
                    max_record_bytes: RECORD_LEN - 1,
                },
            )
            .is_err()
        );
    }

    #[test]
    fn generic_record_validation_rejects_type_length_and_children() {
        let mut record = PowerPointAnimationHash10::new(9).to_record();
        record.record_type_raw ^= 1;
        assert!(PowerPointAnimationHash10::parse_record(&record).is_err());

        let mut record = PowerPointAnimationHash10::new(9).to_record();
        record.data_length = 3;
        assert!(PowerPointAnimationHash10::parse_record(&record).is_err());

        let mut record = PowerPointAnimationHash10::new(9).to_record();
        record
            .children
            .push(PowerPointAnimationHash10::new(1).to_record());
        assert!(PowerPointAnimationHash10::parse_record(&record).is_err());
    }
}
