//! MS-PPT validation for a document container's terminal records.

use super::model::DocumentStructure;
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

impl DocumentStructure {
    /// Validate the exact MS-PPT document structure and its terminal records.
    pub(crate) fn parse(document: &Record) -> Result<Self> {
        super::validation::validate_document(document, super::model::Limits::default())
    }
}

/// Serialize one validated `DocumentContainer` while retaining all opaque
/// records and their source order.
pub(super) fn encode_document(document: &Record) -> Result<Vec<u8>> {
    if document.record_type != RecordType::Document {
        return Err(Error::InvalidFormat(
            "document-structure encoding requires a DocumentContainer root".into(),
        ));
    }
    encode_record(document, true)
}

/// Refresh the wire payloads of the root and its known structural lists after
/// a transaction-local child reorder.
pub(super) fn synchronize(document: &mut Record) -> Result<()> {
    if document.record_type != RecordType::Document {
        return Err(Error::InvalidFormat(
            "document-structure synchronization requires a DocumentContainer root".into(),
        ));
    }
    for child in &mut document.children {
        if child.record_type == RecordType::SlideListWithText {
            sync_container(child)?;
        }
    }
    sync_container(document)
}

fn sync_container(record: &mut Record) -> Result<()> {
    let mut payload = Vec::new();
    for child in &record.children {
        payload.extend_from_slice(&encode_record(child, false)?);
    }
    record.data_length = u32::try_from(payload.len())
        .map_err(|_err| Error::InvalidFormat("PPT structural payload exceeds u32::MAX".into()))?;
    record.data = payload;
    Ok(())
}

fn encode_record(record: &Record, force_children: bool) -> Result<Vec<u8>> {
    let payload = if force_children || record.record_type == RecordType::SlideListWithText {
        let mut payload = Vec::new();
        for child in &record.children {
            payload.extend_from_slice(&encode_record(child, false)?);
        }
        payload
    } else if record.children.is_empty() {
        record.data.clone()
    } else {
        let mut payload = Vec::new();
        for child in &record.children {
            payload.extend_from_slice(&encode_record(child, false)?);
        }
        payload
    };
    let payload_len = u32::try_from(payload.len())
        .map_err(|_err| Error::InvalidFormat("PPT record payload exceeds u32::MAX".into()))?;
    if record.version > 0x000f || record.instance > 0x0fff {
        return Err(Error::InvalidFormat(
            "PPT record version or instance is out of range".into(),
        ));
    }
    let version_instance = record.version | (record.instance << 4);
    let mut bytes = Vec::with_capacity(8usize.saturating_add(payload.len()));
    bytes.extend_from_slice(&version_instance.to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&payload_len.to_le_bytes());
    bytes.extend_from_slice(&payload);
    Ok(bytes)
}
