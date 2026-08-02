//! PowerPoint 9 envelope state from MS-PPT 2.11.3.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;

const ENVELOPE_FLAGS_RECORD_TYPE: u16 = 0x1784;
const ENVELOPE_ALLOWED_FLAGS: u32 = 0x13;

/// Strictly validated document-envelope state.
///
/// This is inert metadata. It does not send mail, invoke a mail client, or
/// interpret an `EnvelopeData9Atom` payload.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointEnvelopeSettings {
    pub has_envelope: bool,
    pub visible: bool,
    pub modified_since_last_send: bool,
}

impl PowerPointEnvelopeSettings {
    /// Construct a representable envelope state.
    pub fn new(has_envelope: bool, visible: bool, modified_since_last_send: bool) -> Result<Self> {
        let value = Self {
            has_envelope,
            visible,
            modified_since_last_send,
        };
        value.validate()?;
        Ok(value)
    }

    /// Parse one strict `EnvelopeFlags9Atom`.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type_raw != ENVELOPE_FLAGS_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "EnvelopeFlags9Atom has an invalid record header or size".to_string(),
            ));
        }
        let flags = u32::from_le_bytes(record.data[0..4].try_into().unwrap());
        if flags & !ENVELOPE_ALLOWED_FLAGS != 0 {
            return Err(PptError::Corrupted(
                "EnvelopeFlags9Atom has nonzero reserved bits".to_string(),
            ));
        }
        let value = Self {
            has_envelope: flags & 0x01 != 0,
            visible: flags & 0x02 != 0,
            modified_since_last_send: flags & 0x10 != 0,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn parse_document(document: &PptRecord) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == ENVELOPE_FLAGS_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(PptError::Corrupted(
                "PPT9 document tag contains multiple EnvelopeFlags9Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Encode the exact four-byte flag atom.
    pub fn to_record(self) -> Result<PptRecord> {
        self.validate()?;
        let flags = u32::from(self.has_envelope)
            | u32::from(self.visible) << 1
            | u32::from(self.modified_since_last_send) << 4;
        Ok(PptRecord {
            record_type: PptRecordType::from(ENVELOPE_FLAGS_RECORD_TYPE),
            record_type_raw: ENVELOPE_FLAGS_RECORD_TYPE,
            version: 0,
            instance: 0,
            data_length: 4,
            data: flags.to_le_bytes().to_vec(),
            children: Vec::new(),
        })
    }

    fn validate(self) -> Result<()> {
        if (self.visible || self.modified_since_last_send) && !self.has_envelope {
            return Err(PptError::Corrupted(
                "Visible or modified envelope state requires an envelope".to_string(),
            ));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_valid_envelope_state() {
        let expected = PowerPointEnvelopeSettings::new(true, true, true).unwrap();
        let record = expected.to_record().unwrap();
        assert_eq!(record.data, 0x13u32.to_le_bytes());
        assert_eq!(
            PowerPointEnvelopeSettings::parse(&record).unwrap(),
            expected
        );
    }

    #[test]
    fn rejects_dependencies_and_reserved_bits() {
        assert!(PowerPointEnvelopeSettings::new(false, true, false).is_err());
        assert!(PowerPointEnvelopeSettings::new(false, false, true).is_err());
        let mut record = PowerPointEnvelopeSettings::default().to_record().unwrap();
        record.data.copy_from_slice(&0x04u32.to_le_bytes());
        assert!(PowerPointEnvelopeSettings::parse(&record).is_err());
        record.data.copy_from_slice(&0x02u32.to_le_bytes());
        assert!(PowerPointEnvelopeSettings::parse(&record).is_err());
    }
}
