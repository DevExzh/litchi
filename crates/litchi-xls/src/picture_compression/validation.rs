//! Resource and cardinality validation for detached snapshots.

use crate::{Error, Result};

use super::RECORD_TYPE;
use super::model::{MAX_RECORD_PAYLOAD, MAX_RECORDS, MAX_STREAM_BYTES, Record, Snapshot};

pub(crate) fn validate(snapshot: &Snapshot) -> Result<()> {
    if snapshot.records.len() > MAX_RECORDS {
        return Err(invalid("snapshot record count exceeds 1024"));
    }
    let mut settings = 0usize;
    let mut encoded_size = 0usize;
    for record in &snapshot.records {
        let payload_len = match record {
            Record::Settings(value) => {
                settings += 1;
                let payload_len = value.payload_len();
                if payload_len > MAX_RECORD_PAYLOAD {
                    return Err(invalid("CompressPictures payload exceeds 8224 bytes"));
                }
                payload_len
            },
            Record::Unknown(value) => {
                if value.payload().len() > MAX_RECORD_PAYLOAD {
                    return Err(invalid("unknown BIFF payload exceeds 8224 bytes"));
                }
                value.payload().len()
            },
        };
        encoded_size = encoded_size
            .checked_add(4)
            .and_then(|size| size.checked_add(payload_len))
            .ok_or_else(|| invalid("snapshot encoded size overflows"))?;
        if encoded_size > MAX_STREAM_BYTES {
            return Err(invalid("snapshot encoded size exceeds 1 MiB"));
        }
    }
    if settings > 1 {
        return Err(invalid(
            "snapshot contains more than one CompressPictures record",
        ));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
