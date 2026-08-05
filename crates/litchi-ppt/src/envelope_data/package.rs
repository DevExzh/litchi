use super::codec::ENVELOPE_DATA_RECORD_TYPE;
use super::model::EnvelopeData;
use crate::package::{Error, Result};
use crate::records::Record;

impl EnvelopeData {
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == ENVELOPE_DATA_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT9 document tag contains multiple EnvelopeData9Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }
}
