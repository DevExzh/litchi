//! PowerPoint 10 document privacy metadata from MS-PPT 2.4.8.

use crate::consts::RecordType;

use super::package::{Error, Result};
use super::records::Record;

const FILTER_PRIVACY_FLAGS_RECORD_TYPE: u16 = 0x36b0;

/// Privacy preferences stored in `FilterPrivacyFlags10Atom`.
///
/// Reading this flag has no side effects. In particular, the library does not
/// remove metadata merely because the preference is enabled.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PrivacySettings {
    /// Remove personally identifiable information when the producer saves.
    pub remove_personally_identifiable_information_on_save: bool,
}

impl PrivacySettings {
    /// Parse one strict `FilterPrivacyFlags10Atom`.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type_raw != FILTER_PRIVACY_FLAGS_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(Error::Corrupted(
                "FilterPrivacyFlags10Atom has an invalid record header or size".to_string(),
            ));
        }
        let flags = u32::from_le_bytes(record.data[0..4].try_into().unwrap());
        if flags & !1 != 0 {
            return Err(Error::Corrupted(
                "FilterPrivacyFlags10Atom has nonzero reserved bits".to_string(),
            ));
        }
        Ok(Self {
            remove_personally_identifiable_information_on_save: flags != 0,
        })
    }

    /// Discover the single privacy atom in the PPT10 document tag.
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(10)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == FILTER_PRIVACY_FLAGS_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT10 document tag contains multiple FilterPrivacyFlags10Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Encode the exact PowerPoint 10 atom.
    pub fn to_record(self) -> Record {
        let flags = u32::from(self.remove_personally_identifiable_information_on_save);
        Record {
            record_type: RecordType::from(FILTER_PRIVACY_FLAGS_RECORD_TYPE),
            record_type_raw: FILTER_PRIVACY_FLAGS_RECORD_TYPE,
            version: 0,
            instance: 0,
            data_length: 4,
            data: flags.to_le_bytes().to_vec(),
            children: Vec::new(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_both_privacy_states() {
        for enabled in [false, true] {
            let expected = PrivacySettings {
                remove_personally_identifiable_information_on_save: enabled,
            };
            let record = expected.to_record();
            assert_eq!(record.record_type_raw, 0x36b0);
            assert_eq!(PrivacySettings::parse(&record).unwrap(), expected);
        }
    }

    #[test]
    fn rejects_reserved_bits_and_invalid_headers() {
        let mut record = PrivacySettings::default().to_record();
        record.data.copy_from_slice(&2u32.to_le_bytes());
        assert!(PrivacySettings::parse(&record).is_err());
        record.data.copy_from_slice(&0u32.to_le_bytes());
        record.instance = 1;
        assert!(PrivacySettings::parse(&record).is_err());
    }
}
